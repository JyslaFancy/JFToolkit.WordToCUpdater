using System.Runtime.InteropServices;

namespace JFToolkit.WordToCUpdater.ComInterop;

/// <summary>
/// Late-bound Word.Application wrapper. No compile-time dependency on any
/// Word interop assembly — works with Word 2007 through Office 365.
/// 
/// Manages the full COM lifecycle: creation → use → cleanup.
/// Implements IDisposable so callers can use `using`.
/// </summary>
internal sealed class WordApplication : IDisposable
{
    // RCWs we hold and must release
    private object? _word;
    private object? _docs;
    private bool _disposed;

    /// <summary>
    /// The Word.Application COM object. Late-bound (dynamic).
    /// </summary>
    public dynamic ComObject => _word!;

    /// <summary>
    /// Create a new Word.Application instance. Hidden, no alerts.
    /// Throws if Word is not installed.
    /// </summary>
    public static WordApplication Create()
    {
        var type = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Word is not installed. " +
                "Install Word 2007 or later, or Office 365.");

        var word = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Failed to create Word.Application instance.");

        dynamic w = word;

        try
        {
            w.Visible = false;            // Don't flash Word window
            w.DisplayAlerts = 0;         // WdAlertLevel.wdAlertsNone
            w.ScreenUpdating = false;     // Performance: don't paint
            w.AutomationSecurity = 3;     // msoAutomationSecurityForceDisable (disable macros for safety)

            // Disable the startup prompt about safe mode after crashes
            try { w.Options.AllowReadingMode = false; } catch { /* not all versions support this */ }
        }
        catch
        {
            // If configuration fails, release what we have and bail
            try { w.Quit(0, 0, 0); } catch { /* best effort */ }
            Marshal.ReleaseComObject(word);
            throw;
        }

        return new WordApplication { _word = word };
    }

    /// <summary>
    /// Open a document. Returns a WordDocument wrapper that must be disposed.
    /// </summary>
    /// <remarks>
    /// <para><b>Password-protected documents:</b> <c>PasswordDocument: ""</c> tells Word
    /// that no password is available. Encrypted documents throw immediately
    /// instead of showing a password dialog that would block the STA thread
    /// (because <c>DisplayAlerts = 0</c> does not suppress the built-in
    /// password prompt, which is a different dialog class).</para>
    /// </remarks>
    public WordDocument OpenDocument(string path)
    {
        ThrowIfDisposed();

        _docs ??= ((dynamic)_word!).Documents;

        dynamic docs = _docs;
        dynamic doc;

        try
        {
            // PasswordDocument: "" → fail fast on encrypted docs instead of
            // blocking the STA thread with a password dialog that never resolves
            doc = docs.Open(
                FileName: path,
                ReadOnly: false,
                PasswordDocument: "",
                Visible: false);
        }
        catch (Exception ex)
        {
            var message = $"Failed to open '{path}'.";
            if (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase) ||
                ex.Message.Contains("encrypt", StringComparison.OrdinalIgnoreCase) ||
                ex.Message.Contains("protected", StringComparison.OrdinalIgnoreCase))
            {
                message += " The document may be password-protected or encrypted.";
            }
            else
            {
                message += " Is the file locked or does it exist?";
            }
            throw new IOException(message, ex);
        }

        return new WordDocument(doc);
    }

    /// <summary>
    /// Called by WordDocument when the document is closed.
    /// Tracks which documents are alive so Quit works correctly.
    /// </summary>
    internal void NotifyDocumentClosed()
    {
        // In a more sophisticated version, track open-count.
        // For now, the caller manages document lifetime through WordDocument.Dispose.
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            if (_word is not null)
            {
                dynamic w = _word;
                // Quit without saving changes to any remaining docs
                w.Quit(SaveChanges: 0, OriginalFormat: 0, RouteDocument: 0);
            }
        }
        catch
        {
            // Word may already be shutting down — ignore
        }

        ComHelper.Release(ref _docs);
        ComHelper.Release(ref _word);

        // Force collection of RCWs so Word.exe actually exits
        ComHelper.CollectFinalizers();
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(WordApplication));
    }
}
