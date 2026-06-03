# JFToolkit.WordToCUpdater

**Zero-dependency NuGet that updates Word Table of Contents page numbers.
Works with Word 2007 through Office 365. No version coupling.**

[![NuGet](https://img.shields.io/badge/nuget-v0.1.0-blue)](https://www.nuget.org/packages/JFToolkit.WordToCUpdater)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Target](https://img.shields.io/badge/targets-.NET%208%20%7C%20.NET%209-512bd4)]()

## The problem

You have a Word document with a table of contents. You add 20 pages of content.
The page numbers in the TOC are now wrong. You have to:

1. Open Word
2. Right-click the TOC
3. Click "Update Field"
4. Choose "Update page numbers only"
5. Save and close

Now imagine doing this for 50 documents. Or doing it from a C# service that
generates Word reports. This library makes it one line of code.

## Install

```bash
dotnet add package JFToolkit.WordToCUpdater
```

No other dependencies. No Office PIA references. No SDK requirements beyond .NET 8+.

## Usage

### One-liner

```csharp
using JFToolkit.WordToCUpdater;

var result = JFToolkit.WordToCUpdater.Update(@"C:\Reports\Q2-Report.docx");

Console.WriteLine(result.Success
    ? $"Done. {result.TocsUpdated} TOCs updated. {result.Elapsed.TotalSeconds:F1}s"
    : $"Failed: {result.Error}");
```

### With options

```csharp
var result = JFToolkit.WordToCUpdater.Update(@"C:\Reports\spec.docx", new TocUpdateOptions
{
    UpdatePageNumbersOnly = false,  // full rebuild — use if headings changed
    CreateBackup = true,            // leaves spec.docx.bak
    OutputPath = @"C:\Reports\spec_final.docx"  // don't touch original
});

Console.WriteLine($"Pages: {result.PagesBefore} → {result.PagesAfter}");
Console.WriteLine($"Word version: {result.WordVersion}");
```

### Async queue (update many documents without blocking)

```csharp
await using var queue = new TocUpdateQueue();

var tasks = Directory.GetFiles(@"C:\Reports", "*.docx")
    .Select(path => queue.EnqueueAsync(path));

var results = await Task.WhenAll(tasks);

foreach (var r in results)
    Console.WriteLine($"{r.OutputPath}: {(r.Success ? "✓" : "✗")} {r.Elapsed.TotalSeconds:F1}s");
```

### Check if Word is installed

```csharp
var version = JFToolkit.WordToCUpdater.DetectVersion();
// "Office365", "Word2016", "Word2013", "NotInstalled", etc.
```

## How it works

Late-bound COM interop via `dynamic`. No compile-time reference to any Word
interop assembly — the same binary works with Word 2007, 2010, 2013, 2016,
2019, and Office 365.

The `TocUpdateQueue` spins up a dedicated STA thread with a Windows message
pump — COM requires this. Your code stays async, the COM calls happen on the
right thread, and Word.Application is created and destroyed cleanly (no zombie
WINWORD.EXE processes).

## Requirements

- **Windows** (Word COM is Windows-only)
- **Microsoft Word** installed (any version 2007+)
- **.NET 8 or .NET 9**

## Error cases handled

| Scenario | Behaviour |
|----------|-----------|
| Word not installed | Returns `Success=false` with clear message |
| File not found | Returns `Success=false` before touching COM |
| File locked by another process | COM error caught, reported in `Error` |
| Document has no TOC | Returns `TocsUpdated=0`, `Success=true` (no-op) |
| Word.exe hangs | `Dispose` force-quits after 5s timeout on the STA thread |
| Multiple concurrent updates | Queued sequentially on one STA thread — no conflicts |

## Version compatibility

Late binding means this library ships zero interop assemblies. It queries
`Type.GetTypeFromProgID("Word.Application")` at runtime. The COM API for
`TablesOfContents.Update()` has been stable since Word 97. This library
targets the intersection that works across all versions.

The `WordVersion` enum is informational only — no behaviour branches on version.

## What it does NOT do

- Does not create or format TOC fields (use Open XML SDK for document creation)
- Does not render Word → PDF
- Does not run on Linux/macOS (Word COM is Windows-only)
- Does not handle password-protected documents

## Building from source

```bash
git clone https://github.com/you/JFToolkit.WordToCUpdater.git
cd JFToolkit.WordToCUpdater
dotnet build
dotnet test
```

## License

MIT — use it anywhere, commercial or personal.
