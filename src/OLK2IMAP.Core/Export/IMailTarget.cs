using MimeKit;

namespace OLK2IMAP.Core.Export;

/// <summary>
/// A destination mail store the extracted messages are written to
/// (IMAP, MBOX, Exchange/EWS, PST). Implementations preserve the folder path
/// and the read/draft state of each message.
/// </summary>
public interface IMailTarget : IAsyncDisposable
{
    string Name { get; }

    /// <summary>Open the connection / output. Called once before any append.</summary>
    Task ConnectAsync(CancellationToken ct);

    /// <summary>Create the folder (and any missing parents) if it does not exist.</summary>
    Task EnsureFolderAsync(string folderPath, CancellationToken ct);

    /// <summary>Append one message into the given folder path.</summary>
    Task AppendAsync(string folderPath, MimeMessage message, bool seen, bool draft, CancellationToken ct);

    /// <summary>Flush / finalize. Called once after all appends.</summary>
    Task FinishAsync(CancellationToken ct);
}
