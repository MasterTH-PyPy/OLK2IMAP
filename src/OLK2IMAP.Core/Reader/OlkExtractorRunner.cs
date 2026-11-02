using System.Diagnostics;

namespace OLK2IMAP.Core.Reader;

/// <summary>
/// Drives the bundled Python extractor (olkextract.exe, a PyInstaller build of
/// extractor/olkextract.py) that reads the Olk IndexedDB and writes an EML tree
/// plus manifest.json. The C# side then imports that into the chosen target.
/// </summary>
public sealed class OlkExtractorRunner
{
    private readonly Action<string> _log;

    public OlkExtractorRunner(Action<string> log) => _log = log;

    /// <summary>Full path to olkextract.exe (bundled next to the app under /extractor),
    /// or null to fall back to "python olkextract.py" for development.</summary>
    public string? ExtractorExePath { get; set; } =
        Path.Combine(AppContext.BaseDirectory, "extractor", "olkextract.exe");

    public string? DevScriptPath { get; set; }

    public async Task<int> RunAsync(string olkProfileDir, string outDir, CancellationToken ct)
    {
        Directory.CreateDirectory(outDir);

        ProcessStartInfo psi;
        if (ExtractorExePath is not null && File.Exists(ExtractorExePath))
        {
            psi = new ProcessStartInfo(ExtractorExePath);
            AddArgs(psi, olkProfileDir, outDir);
        }
        else if (DevScriptPath is not null && File.Exists(DevScriptPath))
        {
            psi = new ProcessStartInfo("python");
            psi.ArgumentList.Add(DevScriptPath);
            AddArgs(psi, olkProfileDir, outDir);
        }
        else
        {
            throw new FileNotFoundException(
                "olkextract.exe nicht gefunden. Erwartet unter: " + ExtractorExePath);
        }

        psi.RedirectStandardError = true;
        psi.RedirectStandardOutput = true;
        psi.UseShellExecute = false;
        psi.CreateNoWindow = true;

        using var p = new Process { StartInfo = psi, EnableRaisingEvents = true };
        p.ErrorDataReceived += (_, e) => { if (e.Data is not null) _log(e.Data); };
        p.OutputDataReceived += (_, e) => { if (e.Data is not null) _log("[extract] " + e.Data); };
        p.Start();
        p.BeginErrorReadLine();
        p.BeginOutputReadLine();
        await p.WaitForExitAsync(ct);
        return p.ExitCode;
    }

    private static void AddArgs(ProcessStartInfo psi, string profile, string outDir)
    {
        psi.ArgumentList.Add("--profile");
        psi.ArgumentList.Add(profile);
        psi.ArgumentList.Add("--out");
        psi.ArgumentList.Add(outDir);
        psi.ArgumentList.Add("--json-progress");
    }
}
