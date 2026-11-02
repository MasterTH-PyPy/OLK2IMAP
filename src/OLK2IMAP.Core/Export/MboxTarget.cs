using System.Text;
using MimeKit;

namespace OLK2IMAP.Core.Export;

/// <summary>
/// Writes one Unix mbox file per folder into an output directory (sub-folders
/// flattened as "Parent__Child.mbox"). Importable into Thunderbird via
/// ImportExportTools NG.
/// </summary>
public sealed class MboxTarget : IMailTarget
{
    private readonly MboxOptions _o;
    private readonly Dictionary<string, FileStream> _streams = new(StringComparer.OrdinalIgnoreCase);

    public MboxTarget(MboxOptions o) => _o = o;

    public string Name => "MBOX";

    public Task ConnectAsync(CancellationToken ct)
    {
        Directory.CreateDirectory(_o.OutDir);
        return Task.CompletedTask;
    }

    public Task EnsureFolderAsync(string folderPath, CancellationToken ct)
    {
        GetStream(folderPath);
        return Task.CompletedTask;
    }

    public async Task AppendAsync(string folderPath, MimeMessage message, bool seen, bool draft, CancellationToken ct)
    {
        var fs = GetStream(folderPath);
        // mbox separator line
        var from = message.From.Mailboxes.FirstOrDefault()?.Address ?? "MAILER-DAEMON";
        var date = (message.Date == default ? DateTimeOffset.UtcNow : message.Date).UtcDateTime
                   .ToString("ddd MMM dd HH:mm:ss yyyy");
        var sep = Encoding.ASCII.GetBytes($"From {from} {date}\n");
        await fs.WriteAsync(sep, ct);
        // write the message; MimeKit's mbox format options escape "From " lines in the body
        var opts = FormatOptions.Default.Clone();
        opts.NewLineFormat = NewLineFormat.Unix;
        using var ms = new MemoryStream();
        await message.WriteToAsync(opts, ms, ct);
        ms.Position = 0;
        await ms.CopyToAsync(fs, ct);
        await fs.WriteAsync(Encoding.ASCII.GetBytes("\n"), ct);
    }

    private FileStream GetStream(string folderPath)
    {
        if (_streams.TryGetValue(folderPath, out var s)) return s;
        var name = (string.IsNullOrEmpty(folderPath) ? "Inbox" : folderPath).Replace('/', '_').Replace("__", "_");
        name = string.Join("__", folderPath.Split('/'));
        foreach (var c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        var path = Path.Combine(_o.OutDir, name + ".mbox");
        var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read);
        _streams[folderPath] = fs;
        return fs;
    }

    public Task FinishAsync(CancellationToken ct) => Task.CompletedTask;

    public async ValueTask DisposeAsync()
    {
        foreach (var s in _streams.Values)
            await s.DisposeAsync();
        _streams.Clear();
    }
}
