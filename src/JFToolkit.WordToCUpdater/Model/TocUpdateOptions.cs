namespace JFToolkit.WordToCUpdater.Model;

/// <summary>
/// Configuration for a TOC update operation.
/// </summary>
public class TocUpdateOptions
{
    /// <summary>
    /// If true, only page numbers are refreshed (fast).
    /// If false, the entire TOC is rebuilt from heading text (slow but catches edited headings).
    /// Default: true.
    /// </summary>
    public bool UpdatePageNumbersOnly { get; set; } = true;

    /// <summary>
    /// If true, all fields in the document are updated (TOC, cross-refs, seq fields, etc.).
    /// If false, only TOC fields are touched.
    /// Default: false.
    /// </summary>
    public bool UpdateAllFields { get; set; } = false;

    /// <summary>
    /// If true, creates a .bak copy of the original file before modifying.
    /// Default: false.
    /// </summary>
    public bool CreateBackup { get; set; } = false;

    /// <summary>
    /// If non-null, saves the result to this path instead of overwriting the original.
    /// </summary>
    public string? OutputPath { get; set; }
}
