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
    /// <remarks>
    /// <para><b>Security note:</b> The backup is written to <c>documentPath + ".bak"</c>
    /// with <c>overwrite: true</c>. The path is derived from the caller-supplied document
    /// path — the caller is responsible for ensuring the path is trusted and
    /// targets the intended directory. No path canonicalization or traversal
    /// checks are performed.</para>
    /// </remarks>
    public bool CreateBackup { get; set; } = false;

    /// <summary>
    /// If non-null, saves the result to this path instead of overwriting the original.
    /// </summary>
    /// <remarks>
    /// <para><b>Security note:</b> This path is passed directly to Word's <c>SaveAs2</c>
    /// without validation. The caller is responsible for ensuring the path is
    /// trusted and resolves to the intended location. Passing user-controlled or
    /// unvalidated input may allow arbitrary file overwrite or directory traversal
    /// (e.g. <c>..\..\etc\critical.dll</c> or absolute system paths).</para>
    /// </remarks>
    public string? OutputPath { get; set; }
}
