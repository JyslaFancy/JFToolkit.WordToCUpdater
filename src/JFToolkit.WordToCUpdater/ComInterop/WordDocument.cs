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
    /// </summary>
    /// <param name="pageNumbersOnly">
    /// If true, only updates page numbers (faster, won't re-read heading text).
    /// If false, fully rebuilds the TOC (slower, catches changed text).
    /// </param>
    /// <returns>Number of TOCs updated.</returns>
    public int UpdateAllTocs(bool pageNumbersOnly = true)
    {
        ThrowIfDisposed();

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
    /// </summary>
    public void UpdateAllFields()
    {
        ThrowIfDisposed();
        _doc.Fields.Update();
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
