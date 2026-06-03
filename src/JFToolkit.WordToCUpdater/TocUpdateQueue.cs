using System.Collections.Concurrent;
using JFToolkit.WordToCUpdater.ComInterop;
using JFToolkit.WordToCUpdater.Model;

namespace JFToolkit.WordToCUpdater;

/// <summary>
/// Queues TOC update operations on a dedicated STA thread with a Windows
/// message pump — required for COM interop to work correctly.
/// 
/// The queue processes one document at a time (Word is single-instance),
/// but exposes an async API so your application doesn't block.
/// </summary>
public sealed class TocUpdateQueue : IDisposable
{
    private readonly Thread _staThread;
    private readonly BlockingCollection<WorkItem> _queue = new();
    private readonly CancellationTokenSource _cts = new();
    private bool _started;
    private bool _disposed;

    public TocUpdateQueue()
    {
        _staThread = new Thread(RunStaPump)
        {
            Name = "WordTocUpdater-STA",
            IsBackground = true,
        };
        _staThread.SetApartmentState(ApartmentState.STA);
    }

    /// <summary>
    /// Enqueue a TOC update. The returned Task completes when the update finishes.
    /// </summary>
    public Task<TocUpdateResult> EnqueueAsync(
        string documentPath, TocUpdateOptions? options = null)
    {
        ThrowIfDisposed();

        options ??= new TocUpdateOptions();

        var tcs = new TaskCompletionSource<TocUpdateResult>();
        _queue.Add(new WorkItem(documentPath, options, tcs));

        // Start the STA thread lazily on first enqueue
        if (!_started)
        {
            _started = true;
            _staThread.Start();
        }

        return tcs.Task;
    }

    private void RunStaPump()
    {
        try
        {
            foreach (var item in _queue.GetConsumingEnumerable(_cts.Token))
            {
                var result = ProcessItem(item);
                item.Completion.SetResult(result);
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown
        }
        catch (Exception ex)
        {
            // If the pump crashes, complete remaining items as failed
            while (_queue.TryTake(out var remaining))
            {
                remaining.Completion.TrySetResult(
                    TocUpdateResult.Failed($"STA pump crashed: {ex.Message}", TimeSpan.Zero));
            }
        }
    }

    private static TocUpdateResult ProcessItem(WorkItem item)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            // Pre-flight checks
            if (!System.IO.File.Exists(item.Path))
            {
                return TocUpdateResult.Failed($"File not found: {item.Path}", sw.Elapsed);
            }

            var version = WordVersionDetector.Detect();
            if (version == WordVersion.NotInstalled)
            {
                return TocUpdateResult.Failed("Microsoft Word is not installed.", sw.Elapsed);
            }

            // Optional backup
            if (item.Options.CreateBackup && !item.Options.OutputPath.IsSet())
            {
                var backup = item.Path + ".bak";
                System.IO.File.Copy(item.Path, backup, overwrite: true);
            }

            using var word = WordApplication.Create();
            using var doc = word.OpenDocument(item.Path);

            var pagesBefore = doc.PageCount;
            var tocCount = doc.TocCount;

            if (item.Options.UpdateAllFields)
            {
                doc.UpdateAllFields();
            }
            else
            {
                doc.UpdateAllTocs(item.Options.UpdatePageNumbersOnly);
            }

            var pagesAfter = doc.PageCount;

            // Save
            if (item.Options.OutputPath.IsSet())
            {
                doc.SaveAs(item.Options.OutputPath!);
            }
            else
            {
                doc.Save();
            }

            return new TocUpdateResult
            {
                Success = true,
                TocsUpdated = tocCount,
                PagesBefore = pagesBefore,
                PagesAfter = pagesAfter,
                WordVersion = version.ToString(),
                OutputPath = item.Options.OutputPath ?? item.Path,
                Elapsed = sw.Elapsed
            };
        }
        catch (Exception ex)
        {
            return TocUpdateResult.Failed(ex.Message, sw.Elapsed);
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _cts.Cancel();
        _queue.CompleteAdding();

        // Give the STA thread 5 seconds to drain and exit
        if (_staThread.IsAlive && !_staThread.Join(5000))
        {
            // Last resort — but COM will clean up when the process exits
        }

        _cts.Dispose();
        _queue.Dispose();
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(TocUpdateQueue));
    }

    private record WorkItem(
        string Path,
        TocUpdateOptions Options,
        TaskCompletionSource<TocUpdateResult> Completion);
}

internal static class StringExtensions
{
    public static bool IsSet(this string? s) => !string.IsNullOrWhiteSpace(s);
}
