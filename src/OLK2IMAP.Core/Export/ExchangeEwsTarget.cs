using Microsoft.Exchange.WebServices.Data;
using MimeKit;
// EWS defines its own "Task" item type in Microsoft.Exchange.WebServices.Data,
// which collides with System.Threading.Tasks.Task. Alias so our async return
// types resolve correctly.
using Task = System.Threading.Tasks.Task;
// "MimeContent" exists in both EWS and MimeKit -> disambiguate to the EWS one.
using MimeContent = Microsoft.Exchange.WebServices.Data.MimeContent;

namespace OLK2IMAP.Core.Export;

/// <summary>
/// Appends messages to an on-prem Exchange mailbox via EWS (Exchange Web
/// Services). Uses the community .NET port of the EWS Managed API.
///
/// NOTE: EWS must be enabled on the Exchange server and the account needs
/// permission to the target mailbox. Modern Exchange may require OAuth instead
/// of basic auth; swap WebCredentials for OAuthCredentials if so.
/// </summary>
public sealed class ExchangeEwsTarget : IMailTarget
{
    private readonly EwsOptions _o;
    private ExchangeService? _svc;
    private readonly Dictionary<string, FolderId> _cache = new(StringComparer.OrdinalIgnoreCase);

    public ExchangeEwsTarget(EwsOptions o) => _o = o;

    public string Name => "Exchange (EWS)";

    public Task ConnectAsync(CancellationToken ct)
    {
        var svc = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
        {
            Credentials = new WebCredentials(_o.User, _o.Password),
        };
        if (_o.AutodiscoverFromUser)
            svc.AutodiscoverUrl(_o.User, _ => true);
        else
            svc.Url = new Uri(_o.EwsUrl);
        _svc = svc;
        return Task.CompletedTask;
    }

    public Task EnsureFolderAsync(string folderPath, CancellationToken ct)
    {
        GetOrCreate(folderPath);
        return Task.CompletedTask;
    }

    public Task AppendAsync(string folderPath, MimeMessage message, bool seen, bool draft, CancellationToken ct)
    {
        var svc = _svc ?? throw new InvalidOperationException("Not connected.");
        var folderId = GetOrCreate(folderPath);

        using var ms = new MemoryStream();
        message.WriteTo(ms);

        var em = new EmailMessage(svc)
        {
            MimeContent = new MimeContent("UTF-8", ms.ToArray()),
        };
        // keep the original sent/received state; file (don't send) into the folder
        em.IsRead = seen;
        var extIsDraft = new ExtendedPropertyDefinition(0x0E07, MapiPropertyType.Integer); // PR_MESSAGE_FLAGS
        // msgflags: 1=READ, 8=UNSENT(draft)
        em.SetExtendedProperty(extIsDraft, (seen ? 1 : 0) | (draft ? 8 : 0));
        em.Save(folderId);
        return Task.CompletedTask;
    }

    private FolderId GetOrCreate(string folderPath)
    {
        var svc = _svc ?? throw new InvalidOperationException("Not connected.");
        if (_cache.TryGetValue(folderPath, out var hit)) return hit;

        FolderId parent = new FolderId(WellKnownFolderName.MsgFolderRoot);
        var builtKey = "";
        foreach (var seg in folderPath.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            builtKey = builtKey.Length == 0 ? seg : builtKey + "/" + seg;
            if (_cache.TryGetValue(builtKey, out var cached)) { parent = cached; continue; }

            var filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, seg);
            var results = svc.FindFolders(parent, filter, new FolderView(1));
            FolderId id;
            if (results.Folders.Count > 0)
            {
                id = results.Folders[0].Id;
            }
            else
            {
                var folder = new Folder(svc) { DisplayName = seg };
                folder.Save(parent);
                id = folder.Id;
            }
            _cache[builtKey] = id;
            parent = id;
        }
        _cache[folderPath] = parent;
        return parent;
    }

    public Task FinishAsync(CancellationToken ct) => Task.CompletedTask;
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;
}
