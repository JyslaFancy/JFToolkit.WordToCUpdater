namespace JFToolkit.WordToCUpdater.Model;

/// <summary>
/// Result of a TOC update operation.
/// </summary>
public class TocUpdateResult
{
    /// <summary>
    /// Whether the operation succeeded.
    /// </summary>
    public bool Success { get; init; }

    /// <summary>
    /// Number of TOC fields found and updated.
    /// </summary>
    public int TocsUpdated { get; init; }

    /// <summary>
    /// Page count before update.
    /// </summary>
    public int PagesBefore { get; init; }

    /// <summary>
    /// Page count after update. May differ from PagesBefore if TOC takes more/less space.
    /// </summary>
    public int PagesAfter { get; init; }

    /// <summary>
    /// Word version detected on the system.
    /// </summary>
    public string WordVersion { get; init; } = "";

    /// <summary>
    /// Path to the output file.
    /// </summary>
    public string OutputPath { get; init; } = "";

    /// <summary>
    /// Time the operation took.
    /// </summary>
    public TimeSpan Elapsed { get; init; }

    /// <summary>
    /// Error message, if any. Null if Success is true.
    /// </summary>
    public string? Error { get; init; }

    public static TocUpdateResult Failed(string error, TimeSpan elapsed)
        => new()
        {
            Success = false,
            Error = error,
            Elapsed = elapsed,
            WordVersion = "Unknown"
        };
}
