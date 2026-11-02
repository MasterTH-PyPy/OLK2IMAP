using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using MimeKit;

namespace OLK2IMAP.Core.Export;

/// <summary>
/// Appends messages to an IMAP server (generic, incl. Exchange/O365 IMAP),
/// (re)creating the folder hierarchy. Read/draft state is set via message flags.
/// </summary>
public sealed class ImapTarget : IMailTarget
{
    private readonly ImapOptions _o;
    private readonly ImapClient _client = new();
    private readonly Dictionary<string, IMailFolder> _cache = new(StringComparer.OrdinalIgnoreCase);
    private char _delim = '/';
    private IMailFolder? _root;

    public ImapTarget(ImapOptions o) => _o = o;

    public string Name => "IMAP";

    public async Task ConnectAsync(CancellationToken ct)
    {
        await _client.ConnectAsync(_o.Host, _o.Port,
            _o.UseSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable, ct);
        await _client.AuthenticateAsync(_o.User, _o.Password, ct);
        _root = _client.GetFolder(_client.PersonalNamespaces[0]);
        _delim = _root.DirectorySeparator == '\0' ? '/' : _root.DirectorySeparator;
    }

    public async Task EnsureFolderAsync(string folderPath, CancellationToken ct)
        => await GetOrCreateAsync(folderPath, ct);

    public async Task AppendAsync(string folderPath, MimeMessage message, bool seen, bool draft, CancellationToken ct)
    {
        var folder = await GetOrCreateAsync(folderPath, ct);
        if (!folder.IsOpen)
            await folder.OpenAsync(FolderAccess.ReadWrite, ct);
        var flags = MessageFlags.None;
        if (seen) flags |= MessageFlags.Seen;
        if (draft) flags |= MessageFlags.Draft;
        var date = message.Date == default ? DateTimeOffset.UtcNow : message.Date;
        await folder.AppendAsync(message, flags, date, ct);
    }

    private async Task<IMailFolder> GetOrCreateAsync(string folderPath, CancellationToken ct)
    {
        if (_root is null) throw new InvalidOperationException("Not connected.");
        var full = string.IsNullOrEmpty(_o.FolderPrefix)
            ? folderPath
            : _o.FolderPrefix.Replace('/', _delim) + _delim + folderPath;
        if (_cache.TryGetValue(full, out var cached)) return cached;

        var parts = full.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var current = _root;
        var builtKey = "";
        foreach (var raw in parts)
        {
            var seg = MapName(raw);
            builtKey = builtKey.Length == 0 ? raw : builtKey + "/" + raw;
            IMailFolder child;
            try
            {
                child = await current.GetSubfolderAsync(seg, ct);
            }
            catch (FolderNotFoundException)
            {
                child = await current.CreateAsync(seg, isMessageFolder: true, ct);
            }
            _cache[builtKey] = child;
            current = child;
        }
        _cache[full] = current;
        return current;
    }

    /// <summary>Map well-known OWA folder names onto standard IMAP equivalents where sensible.</summary>
    private static string MapName(string name) => name switch
    {
        "Sent Items" => "Sent",
        "Deleted Items" => "Trash",
        _ => name,
    };

    public async Task FinishAsync(CancellationToken ct)
    {
        foreach (var f in _cache.Values)
            if (f.IsOpen) await f.CloseAsync(false, ct);
    }

    public async ValueTask DisposeAsync()
    {
        if (_client.IsConnected) await _client.DisconnectAsync(true);
        _client.Dispose();
    }
}
