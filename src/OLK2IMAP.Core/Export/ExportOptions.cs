namespace OLK2IMAP.Core.Export;

/// <summary>Where the extracted mail (EML tree + manifest.json) lives.</summary>
public sealed class ExtractSource
{
    public required string OutDir { get; init; }
    public string ManifestPath => System.IO.Path.Combine(OutDir, "manifest.json");
    public string EmlDir => System.IO.Path.Combine(OutDir, "EML");
}

public sealed class ImapOptions
{
    public string Host { get; set; } = "";
    public int Port { get; set; } = 993;
    public bool UseSsl { get; set; } = true;
    public string User { get; set; } = "";
    public string Password { get; set; } = "";
    /// <summary>Prefix all created folders under this parent (e.g. "Import/Olk"). Optional.</summary>
    public string? FolderPrefix { get; set; }
}

public sealed class EwsOptions
{
    public string EwsUrl { get; set; } = ""; // e.g. https://mail.example.de/EWS/Exchange.asmx
    public string User { get; set; } = "";
    public string Password { get; set; } = "";
    public bool AutodiscoverFromUser { get; set; }
}

public sealed class MboxOptions
{
    public string OutDir { get; set; } = "";
}

public sealed class PstOptions
{
    public string PstPath { get; set; } = "";
}

/// <summary>Global run options.</summary>
public sealed class RunOptions
{
    /// <summary>Do everything except actually writing to the target (count + log only).</summary>
    public bool DryRun { get; set; }
    public bool IncludeDeletedItems { get; set; } = true;
    public bool IncludeDrafts { get; set; } = true;
}
