using System.Runtime.InteropServices;

namespace JFToolkit.WordToCUpdater.ComInterop;

/// <summary>
/// COM lifetime management. The rules:
/// 1. Release in reverse creation order.
/// 2. Never double-dot COM properties (each dot creates an unreleasable RCW).
/// 3. GC.Collect + WaitForPendingFinalizers after ReleaseComObject.
/// </summary>
internal static class ComHelper
{
    /// <summary>
    /// Release a COM object and suppress the finalizer for it.
    /// Safe to call on null references.
    /// </summary>
    public static void Release(ref object? comObject)
    {
        if (comObject == null) return;

        try
        {
            // Only release if it's actually a COM wrapper
            if (Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch
        {
            // The object may already be released or in a bad state.
            // Don't let cleanup failures block the caller.
        }
        finally
        {
            comObject = null;
        }
    }

    /// <summary>
    /// Force GC to collect any unreachable RCWs so Word.exe can exit.
    /// Call after releasing the root COM object (Word.Application).
    /// 
    /// The double-Collect pattern (Collect → WaitForPendingFinalizers → Collect)
    /// is deliberate for COM interop: the first Collect promotes finalizable
    /// RCWs to the finalizer queue; WaitForPendingFinalizers drains that queue;
    /// the second Collect sweeps any objects resurrected during finalization.
    /// Without this, orphaned RCWs can keep WINWORD.EXE alive after Dispose().
    /// </summary>
    public static void CollectFinalizers()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    /// <summary>
    /// Extract an RCW from a dynamic COM return value, assign it to a named variable
    /// so we can explicitly release it later. Use instead of double-dot chaining.
    /// 
    /// Example (wrong):
    ///   dynamic doc = word.Documents.Open(path);  // leaks the Documents RCW
    /// 
    /// Example (correct):
    ///   dynamic docs = word.Documents;      // capture
    ///   dynamic doc = docs.Open(path);      // use captured
    ///   // ... later release docs then doc
    /// </summary>
    public static T Capture<T>(dynamic comObject) where T : class
    {
        return (T)comObject;
    }
}
