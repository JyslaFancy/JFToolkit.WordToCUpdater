using JFToolkit.WordToCUpdater;
using JFToolkit.WordToCUpdater.Model;

// ──────────────────────────────────────────────
// TocUpdater — Console Demo
// ──────────────────────────────────────────────

Console.WriteLine("TocUpdater Demo");
Console.WriteLine("===================");
Console.WriteLine();

// 1. Detect Word version
var version = TocUpdater.DetectVersion();
Console.WriteLine($"Word version detected: {version}");
Console.WriteLine();

if (version == "NotInstalled")
{
    Console.WriteLine("Word is not installed. Exiting.");
    Console.WriteLine("Install Word 2007+ or Office 365 to use this library.");
    return;
}

// 2. Single document — one-liner
Console.WriteLine("--- One-liner update ---");
var result = TocUpdater.Update(@"C:\demo\report.docx");
Console.WriteLine(result.Success
    ? $"✓ {result.TocsUpdated} TOC(s) updated. " +
      $"Pages: {result.PagesBefore} → {result.PagesAfter}. " +
      $"Took {result.Elapsed.TotalSeconds:F1}s"
    : $"✗ Failed: {result.Error}");
Console.WriteLine();

// 3. With options — full rebuild, backup, output to new file
Console.WriteLine("--- Options: full rebuild + backup ---");
var opts = new TocUpdateOptions
{
    UpdatePageNumbersOnly = false,  // Full rebuild from headings
    CreateBackup = true,
    UpdateAllFields = true          // Also update cross-refs, page refs
};
result = TocUpdater.Update(@"C:\demo\contract.docx", opts);
Console.WriteLine(result.Success
    ? $"✓ Full rebuild done. Backup at contract.docx.bak"
    : $"✗ Failed: {result.Error}");
Console.WriteLine();

// 4. Batch processing with async queue
Console.WriteLine("--- Batch: async queue ---");
await BatchUpdateDemo();

Console.WriteLine();
Console.WriteLine("Done. Press any key to exit.");
Console.ReadKey();


static async Task BatchUpdateDemo()
{
    await using var queue = new TocUpdateQueue();

    // Simulate a batch — in real usage, point at actual .docx files
    var paths = new[]
    {
        @"C:\demo\report.docx",
        @"C:\demo\spec.docx",
        @"C:\demo\proposal.docx"
    };

    var tasks = paths.Select(p => queue.EnqueueAsync(p));
    var results = await Task.WhenAll(tasks);

    foreach (var r in results)
    {
        var icon = r.Success ? "✓" : "✗";
        var detail = r.Success
            ? $"{r.TocsUpdated} TOCs, {r.PagesAfter} pages, {r.Elapsed.TotalSeconds:F1}s"
            : r.Error;
        Console.WriteLine($"  {icon} {Path.GetFileName(r.OutputPath)} — {detail}");
    }
}
