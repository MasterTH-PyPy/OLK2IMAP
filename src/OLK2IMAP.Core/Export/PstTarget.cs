using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using MimeKit;

namespace OLK2IMAP.Core.Export;

/// <summary>
/// Writes messages into a Unicode .pst file (no Outlook required) via Aspose.Email.
///
/// NOTE: Aspose.Email is commercial. Without a license it runs in evaluation mode
/// (item-count limits / watermarks). Set the license once at startup via
/// new Aspose.Email.License().SetLicense("Aspose.Email.lic").
/// </summary>
public sealed class PstTarget : IMailTarget
{
    private readonly PstOptions _o;
    private PersonalStorage? _pst;
    private readonly Dictionary<string, FolderInfo> _cache = new(StringComparer.OrdinalIgnoreCase);

    public PstTarget(PstOptions o) => _o = o;

    public string Name => "PST";

    public Task ConnectAsync(CancellationToken ct)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(_o.PstPath))!);
        _pst = PersonalStorage.Create(_o.PstPath, FileFormatVersion.Unicode);
        return Task.CompletedTask;
    }

    public Task EnsureFolderAsync(string folderPath, CancellationToken ct)
    {
        GetOrCreate(folderPath);
        return Task.CompletedTask;
    }

    public Task AppendAsync(string folderPath, MimeMessage message, bool seen, bool draft, CancellationToken ct)
    {
        var folder = GetOrCreate(folderPath);
        using var ms = new MemoryStream();
        message.WriteTo(ms);
        ms.Position = 0;
        var mapi = MapiMessage.Load(ms, new Aspose.Email.EmlLoadOptions());
        if (seen)
            mapi.SetMessageFlags(mapi.Flags | MapiMessageFlags.MSGFLAG_READ);
        if (draft)
            mapi.SetMessageFlags(mapi.Flags | MapiMessageFlags.MSGFLAG_UNSENT);
        folder.AddMessage(mapi);
        return Task.CompletedTask;
    }

    private FolderInfo GetOrCreate(string folderPath)
    {
        var pst = _pst ?? throw new InvalidOperationException("PST not open.");
        if (_cache.TryGetValue(folderPath, out var hit)) return hit;

        FolderInfo current = pst.RootFolder;
        var builtKey = "";
        foreach (var seg in folderPath.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            builtKey = builtKey.Length == 0 ? seg : builtKey + "/" + seg;
            if (_cache.TryGetValue(builtKey, out var cached)) { current = cached; continue; }

            FolderInfo? child = null;
            try { child = current.GetSubFolder(seg); } catch { /* not found */ }
            child ??= current.AddSubFolder(seg);
            _cache[builtKey] = child;
            current = child;
        }
        _cache[folderPath] = current;
        return current;
    }

    public Task FinishAsync(CancellationToken ct) => Task.CompletedTask;

    public ValueTask DisposeAsync()
    {
        _pst?.Dispose();
        return ValueTask.CompletedTask;
    }
}
