using MimeKit;
using OLK2IMAP.Core.Model;

namespace OLK2IMAP.Core.Export;

public sealed record ExportProgress(int Done, int Total, string Folder, string? Note = null);

/// <summary>
/// Orchestrates reading the extracted EML tree + manifest and pushing every
/// message into one target, folder by folder, preserving read/draft state.
/// </summary>
public sealed class Exporter
{
    private readonly ExtractSource _src;
    private readonly RunOptions _opts;
    private readonly Action<string> _log;

    public Exporter(ExtractSource src, RunOptions opts, Action<string> log)
    {
        _src = src;
        _opts = opts;
        _log = log;
    }

    public async Task RunAsync(IMailTarget target, IProgress<ExportProgress>? progress, CancellationToken ct)
    {
        var manifest = ExtractManifest.Load(_src.ManifestPath);

        var messages = manifest.Messages
            .Where(m => _opts.IncludeDrafts || !m.IsDraft)
            .Where(m => _opts.IncludeDeletedItems ||
                        !(m.Folder ?? "").StartsWith("Deleted Items", StringComparison.OrdinalIgnoreCase))
            .ToList();

        int total = messages.Count;
        _log($"{target.Name}: {total} Nachrichten in {messages.Select(m => m.Folder).Distinct().Count()} Ordnern"
             + (_opts.DryRun ? "  [DRY-RUN]" : ""));

        await target.ConnectAsync(ct);

        int done = 0;
        foreach (var group in messages.GroupBy(m => m.Folder ?? "_Unknown")
                                      .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
        {
            ct.ThrowIfCancellationRequested();
            string folder = group.Key;
            if (!_opts.DryRun)
                await target.EnsureFolderAsync(folder, ct);

            foreach (var entry in group.OrderBy(m => m.Date, StringComparer.Ordinal))
            {
                ct.ThrowIfCancellationRequested();
                var emlPath = Path.Combine(_src.EmlDir, entry.File.Replace('/', Path.DirectorySeparatorChar));
                try
                {
                    if (_opts.DryRun)
                    {
                        _log($"  [dry] {folder}  <-  {Path.GetFileName(emlPath)}");
                    }
                    else
                    {
                        await using var fs = File.OpenRead(emlPath);
                        var msg = await MimeMessage.LoadAsync(fs, ct);
                        await target.AppendAsync(folder, msg, entry.IsRead, entry.IsDraft, ct);
                    }
                }
                catch (Exception ex)
                {
                    _log($"  FEHLER {Path.GetFileName(emlPath)}: {ex.Message}");
                }
                done++;
                progress?.Report(new ExportProgress(done, total, folder));
            }
        }

        if (!_opts.DryRun)
            await target.FinishAsync(ct);
        _log($"{target.Name}: fertig ({done}/{total}).");
    }
}
