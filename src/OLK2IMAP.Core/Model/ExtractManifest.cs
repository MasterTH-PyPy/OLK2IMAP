using System.Text.Json;
using System.Text.Json.Serialization;

namespace OLK2IMAP.Core.Model;

/// <summary>
/// manifest.json produced by the bundled olkextract extractor: the folder tree
/// plus one entry per extracted message (pointing at an .eml file under EML/).
/// </summary>
public sealed class ExtractManifest
{
    [JsonPropertyName("account")] public string? Account { get; set; }
    [JsonPropertyName("generated")] public string? Generated { get; set; }
    [JsonPropertyName("folders")] public List<FolderEntry> Folders { get; set; } = new();
    [JsonPropertyName("messages")] public List<MessageEntry> Messages { get; set; } = new();

    private static readonly JsonSerializerOptions Opts = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    public static ExtractManifest Load(string manifestPath)
    {
        using var fs = File.OpenRead(manifestPath);
        return JsonSerializer.Deserialize<ExtractManifest>(fs, Opts)
               ?? throw new InvalidDataException($"Could not parse manifest: {manifestPath}");
    }

    /// <summary>Distinct folder paths that actually contain messages, in a stable order.</summary>
    public IEnumerable<string> MessageFolderPaths()
        => Messages.Select(m => m.Folder ?? "_Unknown")
                   .Distinct()
                   .OrderBy(p => p, StringComparer.OrdinalIgnoreCase);
}

public sealed class FolderEntry
{
    [JsonPropertyName("id")] public string? Id { get; set; }
    [JsonPropertyName("name")] public string? Name { get; set; }
    [JsonPropertyName("path")] public string? Path { get; set; }
    [JsonPropertyName("serverCount")] public int? ServerCount { get; set; }
}

public sealed class MessageEntry
{
    /// <summary>Path of the .eml file relative to the EML/ directory (forward slashes).</summary>
    [JsonPropertyName("file")] public string File { get; set; } = "";
    [JsonPropertyName("folder")] public string? Folder { get; set; }
    [JsonPropertyName("internetMessageId")] public string? InternetMessageId { get; set; }
    [JsonPropertyName("date")] public string? Date { get; set; }
    [JsonPropertyName("isRead")] public bool IsRead { get; set; }
    [JsonPropertyName("isDraft")] public bool IsDraft { get; set; }
    [JsonPropertyName("hasAttachments")] public bool HasAttachments { get; set; }
}
