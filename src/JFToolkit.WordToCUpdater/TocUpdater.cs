using JFToolkit.WordToCUpdater.Model;

namespace JFToolkit.WordToCUpdater;

/// <summary>
/// Synchronous one-shot API. For async usage, use <see cref="TocUpdateQueue"/> instead.
/// 
/// Usage:
///   var result = TocUpdater.Update("C:\\docs\\report.docx");
///   if (result.Success)
///       Console.WriteLine($"Updated {result.TocsUpdated} TOCs. {result.PagesAfter} pages.");
///       
/// <para><b>Security:</b> <paramref name="documentPath"/> and
/// <see cref="TocUpdateOptions.OutputPath"/> are passed to the file system and Word COM
/// without path validation. The caller is responsible for ensuring
/// these values are trusted and resolve to intended locations.</para>
/// </summary>
public static class TocUpdater
{
    /// <summary>
    /// Update all TOC fields in a Word document to correct page numbers.
    /// Blocks the calling thread until complete.
    /// </summary>
    /// <param name="documentPath">Full path to the .docx file.</param>
    /// <param name="options">Optional configuration.</param>
    /// <returns>Result with page counts, timing, and any errors.</returns>
    public static TocUpdateResult Update(string documentPath, TocUpdateOptions? options = null)
    {
        using var queue = new TocUpdateQueue();
        return queue.EnqueueAsync(documentPath, options).GetAwaiter().GetResult();
    }

    /// <summary>
    /// Detect which Word version is installed without doing any document work.
    /// </summary>
    public static string DetectVersion()
        => ComInterop.WordVersionDetector.Detect().ToString();
}
