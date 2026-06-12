namespace JFToolkit.WordToCUpdater.ComInterop;

/// <summary>
/// Late-bound Word Document wrapper. Implements IDisposable so Close()
/// is always called regardless of exceptions.
/// </summary>
internal sealed class WordDocument : IDisposable
{
    private readonly dynamic _doc;
    private bool _disposed;

    internal WordDocument(dynamic doc)
    {
        _doc = doc ?? throw new ArgumentNullException(nameof(doc));
    }

    /// <summary>
    /// Number of tables of contents in this document.
    /// </summary>
    public int TocCount
    {
        get
        {
            ThrowIfDisposed();
            return _doc.TablesOfContents.Count;
        }
    }

    /// <summary>
    /// The full path of this document.
    /// </summary>
    public string Path
    {
        get
        {
            ThrowIfDisposed();
            return _doc.FullName;
        }
    }

    /// <summary>
    /// Number of pages in the document (after TOC update, this changes).
    /// </summary>
    public int PageCount
    {
        get
        {
            ThrowIfDisposed();
            return _doc.ComputeStatistics(2); // wdStatisticPages = 2
        }
    }

    /// <summary>
    /// Update all table-of-contents fields in the document.
    /// Security: Unlinks dangerous fields (INCLUDETEXT/INCLUDEPICTURE/LINK pointing
    /// at remote content) before updating, to prevent SSRF and NTLM hash leakage.
    /// </summary>
    /// <param name="pageNumbersOnly">
    /// If true, only updates page numbers (faster, won't re-read heading text).
    /// If false, fully rebuilds the TOC (slower, catches changed text).
    /// </param>
    /// <returns>Number of TOCs updated.</returns>
    public int UpdateAllTocs(bool pageNumbersOnly = true)
    {
        ThrowIfDisposed();

        // Security: unlink dangerous fields before any field refresh
        UnlinkDangerousFields();

        var count = 0;
        var tocs = _doc.TablesOfContents;
        var tocCount = tocs.Count;

        for (int i = 1; i <= tocCount; i++)
        {
            try
            {
                // WdTocFormat.wdTOCTemplate = 1   (use document's TOC style)
                // WdTocFormat.wdTOCClassic = 2
                tocs[i].Update();
                count++;
            }
            catch
            {
                // Some TOC entries may be broken, skip them
            }
        }

        return count;
    }

    /// <summary>
    /// Update all fields in the document (TOC, cross-references, page numbers, etc.).
    /// Security: Unlinks dangerous fields (INCLUDETEXT/INCLUDEPICTURE/LINK pointing
    /// at remote content) before calling Fields.Update(), to prevent SSRF and NTLM
    /// hash leakage from malicious .docx files.
    /// </summary>
    public void UpdateAllFields()
    {
        ThrowIfDisposed();

        // Security: unlink dangerous fields before bulk-updating
        UnlinkDangerousFields();

        _doc.Fields.Update();
    }

    /// <summary>
    /// Unlink fields that could trigger outbound network connections when refreshed.
    /// Targets INCLUDETEXT, INCLUDEPICTURE, and LINK field types that reference
    /// remote content (UNC paths like \\server\share → NTLM hash leak,
    /// or http(s):// URLs → SSRF).
    /// 
    /// With DisplayAlerts = 0 (set in WordApplication.Create), Word's usual
    /// "update links?" warning is suppressed, so this happens silently unless
    /// we proactively strip dangerous fields first.
    /// </summary>
    private void UnlinkDangerousFields()
    {
        var fields = _doc.Fields;
        var count = fields.Count;

        // Iterate backwards — Unlink() removes the field from the collection,
        // so forward iteration would skip items after a removal
        for (int i = count; i >= 1; i--)
        {
            try
            {
                var field = fields[i];
                int fieldType = (int)field.Type;

                // WdFieldType values:
                //   wdFieldIncludeText    = 46  (INCLUDETEXT)
                //   wdFieldIncludePicture = 50  (INCLUDEPICTURE)
                //   wdFieldLink           = 56  (LINK)
                if (fieldType == 46 || fieldType == 50 || fieldType == 56)
                {
                    var code = field.Code?.Text?.ToString() ?? "";
                    if (IsRemoteFieldCode(code))
                    {
                        field.Unlink();
                    }
                }
            }
            catch
            {
                // Field may be locked, broken, or inaccessible — skip
            }
        }
    }

    /// <summary>
    /// Detect field codes that would cause Word to make outbound network
    /// connections when the field is updated.
    /// 
    /// UNC paths (\\server\share)  → Windows attempts NTLM authentication,
    ///                                leaking the user's credential hash.
    /// HTTP(S) URLs                 → Word performs an HTTP request (SSRF).
    /// </summary>
    private static bool IsRemoteFieldCode(string code)
    {
        if (string.IsNullOrWhiteSpace(code))
            return false;

        // UNC paths: \\server\share — triggers NTLM auth, leaks hash
        // Exclude localhost references (no credential leak)
        if (code.Contains(@"\\") &&
            !code.Contains(@"\\localhost") &&
            !code.Contains(@"\\127.0.0.1"))
        {
            return true;
        }

        // Remote URLs: http://, https:// — SSRF
        if (code.Contains("http://", StringComparison.OrdinalIgnoreCase) ||
            code.Contains("https://", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    /// <summary>
    /// Save the document.
    /// </summary>
    public void Save()
    {
        ThrowIfDisposed();
        _doc.Save();
    }

    /// <summary>
    /// Save the document to a new path.
    /// </summary>
    public void SaveAs(string newPath)
    {
        ThrowIfDisposed();
        _doc.SaveAs2(FileName: newPath);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            _doc.Close(SaveChanges: 0); // wdDoNotSaveChanges
        }
        catch
        {
            // Document may already be closed
        }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(WordDocument));
    }
}
