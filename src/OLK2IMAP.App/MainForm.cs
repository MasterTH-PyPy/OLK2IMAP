using OLK2IMAP.Core.Export;
using OLK2IMAP.Core.Reader;
using OLK2IMAP.Core.Util;

namespace OLK2IMAP.App;

/// <summary>
/// Main window. Standard grey WinForms UI. Two steps:
///   1) Extrahieren  - run the bundled extractor over an Olk profile -> EML + manifest
///   2) Übertragen   - push the extracted mail into IMAP / MBOX / Exchange(EWS) / PST
/// </summary>
public sealed class MainForm : Form
{
    private readonly TextBox _olkPath = new() { Width = 420 };
    private readonly TextBox _workDir = new() { Width = 420 };
    private readonly ComboBox _target = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
    private readonly CheckBox _dryRun = new() { Text = "Dry-Run (nur simulieren)", Checked = true, AutoSize = true };
    private readonly CheckBox _inclDeleted = new() { Text = "Gelöschte einbeziehen", Checked = true, AutoSize = true };
    private readonly CheckBox _inclDrafts = new() { Text = "Entwürfe einbeziehen", Checked = true, AutoSize = true };
    private readonly Button _btnExtract = new() { Text = "1. Extrahieren", Width = 160, Height = 30, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    private readonly Button _btnRun = new() { Text = "2. Übertragen", Width = 160, Height = 30, AutoSizeMode = AutoSizeMode.GrowAndShrink };
    private readonly ProgressBar _progress = new() { Width = 560, Style = ProgressBarStyle.Continuous };
    private readonly TextBox _log = new() { Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Width = 720, Height = 220, Font = new Font("Consolas", 8.5f) };

    // target settings
    private readonly GroupBox _gImap = new() { Text = "IMAP", Width = 720, Height = 92 };
    private readonly TextBox _imHost = new() { Width = 220 };
    private readonly NumericUpDown _imPort = new() { Minimum = 1, Maximum = 65535, Value = 993, Width = 70 };
    private readonly CheckBox _imSsl = new() { Text = "SSL", Checked = true, AutoSize = true };
    private readonly TextBox _imUser = new() { Width = 220 };
    private readonly TextBox _imPass = new() { Width = 180, UseSystemPasswordChar = true };
    private readonly TextBox _imPrefix = new() { Width = 160 };

    private readonly GroupBox _gEws = new() { Text = "Exchange (EWS)", Width = 720, Height = 92, Visible = false };
    private readonly TextBox _ewUrl = new() { Width = 320 };
    private readonly TextBox _ewUser = new() { Width = 220 };
    private readonly TextBox _ewPass = new() { Width = 180, UseSystemPasswordChar = true };

    private readonly GroupBox _gMbox = new() { Text = "MBOX", Width = 720, Height = 60, Visible = false };
    private readonly TextBox _mbDir = new() { Width = 420 };

    private readonly GroupBox _gPst = new() { Text = "PST", Width = 720, Height = 60, Visible = false };
    private readonly TextBox _pstPath = new() { Width = 420 };

    private CancellationTokenSource? _cts;

    public MainForm()
    {
        Text = AppVersion.Display + "  -  new Outlook (Olk) -> IMAP / Exchange / MBOX / PST";
        try { Icon = Icon.ExtractAssociatedIcon(Environment.ProcessPath ?? Application.ExecutablePath); } catch { /* ignore */ }
        Width = 770; Height = 760; StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(770, 760);

        _target.Items.AddRange(new object[] { "IMAP", "MBOX", "Exchange (EWS)", "PST" });
        _target.SelectedIndex = 0;
        _target.SelectedIndexChanged += (_, _) => UpdateTargetPanels();

        BuildLayout();
        UpdateTargetPanels();

        _btnExtract.Click += async (_, _) => await OnExtractAsync();
        _btnRun.Click += async (_, _) => await OnRunAsync();
    }

    private void BuildLayout()
    {
        int y = 12;
        Label L(string t, int yy) { var l = new Label { Text = t, AutoSize = true, Left = 12, Top = yy + 3 }; Controls.Add(l); return l; }

        L("Olk-Profil:", y);
        _olkPath.Left = 110; _olkPath.Top = y; Controls.Add(_olkPath);
        AddBrowseFolder(_olkPath, 540, y);
        y += 30;

        L("Arbeitsordner:", y);
        _workDir.Left = 110; _workDir.Top = y; Controls.Add(_workDir);
        AddBrowseFolder(_workDir, 540, y);
        y += 30;

        _btnExtract.Left = 110; _btnExtract.Top = y; Controls.Add(_btnExtract);
        var hint = new Label { Text = "liest den Offline-Store -> EML + manifest.json", AutoSize = true, Left = 285, Top = y + 7, ForeColor = SystemColors.GrayText };
        Controls.Add(hint);
        y += 46;

        L("Ziel:", y);
        _target.Left = 110; _target.Top = y; Controls.Add(_target);
        y += 34;

        // target panels (stacked at same spot; only one visible)
        BuildImap(); BuildEws(); BuildMbox(); BuildPst();
        foreach (var g in new Control[] { _gImap, _gEws, _gMbox, _gPst }) { g.Left = 12; g.Top = y; Controls.Add(g); }
        y += 100;

        _dryRun.Left = 12; _dryRun.Top = y; Controls.Add(_dryRun);
        _inclDeleted.Left = 200; _inclDeleted.Top = y; Controls.Add(_inclDeleted);
        _inclDrafts.Left = 380; _inclDrafts.Top = y; Controls.Add(_inclDrafts);
        y += 30;

        _btnRun.Left = 12; _btnRun.Top = y; Controls.Add(_btnRun);
        y += 40;

        _progress.Left = 12; _progress.Top = y; Controls.Add(_progress);
        y += 30;

        _log.Left = 12; _log.Top = y; Controls.Add(_log);
    }

    private void BuildImap()
    {
        void A(Control c, int x, int yy) { c.Left = x; c.Top = yy; _gImap.Controls.Add(c); }
        A(new Label { Text = "Host", AutoSize = true }, 10, 22); A(_imHost, 50, 19);
        A(new Label { Text = "Port", AutoSize = true }, 290, 22); A(_imPort, 325, 19);
        A(_imSsl, 410, 20);
        A(new Label { Text = "User", AutoSize = true }, 10, 56); A(_imUser, 50, 53);
        A(new Label { Text = "Pass", AutoSize = true }, 290, 56); A(_imPass, 325, 53);
        A(new Label { Text = "Prefix", AutoSize = true }, 520, 56); A(_imPrefix, 560, 53);
    }

    private void BuildEws()
    {
        void A(Control c, int x, int yy) { c.Left = x; c.Top = yy; _gEws.Controls.Add(c); }
        A(new Label { Text = "EWS-URL", AutoSize = true }, 10, 22); A(_ewUrl, 80, 19);
        A(new Label { Text = "User", AutoSize = true }, 10, 56); A(_ewUser, 80, 53);
        A(new Label { Text = "Pass", AutoSize = true }, 320, 56); A(_ewPass, 360, 53);
    }

    private void BuildMbox()
    {
        var l = new Label { Text = "Ausgabeordner", AutoSize = true, Left = 10, Top = 25 };
        _gMbox.Controls.Add(l); _mbDir.Left = 110; _mbDir.Top = 22; _gMbox.Controls.Add(_mbDir);
        AddBrowseFolderTo(_gMbox, _mbDir, 540, 21);
    }

    private void BuildPst()
    {
        var l = new Label { Text = "PST-Datei", AutoSize = true, Left = 10, Top = 25 };
        _gPst.Controls.Add(l); _pstPath.Left = 110; _pstPath.Top = 22; _gPst.Controls.Add(_pstPath);
        var b = new Button { Text = "...", Left = 540, Top = 20, Width = 30 };
        b.Click += (_, _) => { using var d = new SaveFileDialog { Filter = "PST|*.pst", FileName = "olk-export.pst" }; if (d.ShowDialog() == DialogResult.OK) _pstPath.Text = d.FileName; };
        _gPst.Controls.Add(b);
    }

    private void UpdateTargetPanels()
    {
        _gImap.Visible = _target.SelectedIndex == 0;
        _gMbox.Visible = _target.SelectedIndex == 1;
        _gEws.Visible = _target.SelectedIndex == 2;
        _gPst.Visible = _target.SelectedIndex == 3;
    }

    private void AddBrowseFolder(TextBox tb, int x, int y) => AddBrowseFolderTo(this, tb, x, y);

    private void AddBrowseFolderTo(Control parent, TextBox tb, int x, int y)
    {
        var b = new Button { Text = "...", Left = x, Top = y - 1, Width = 30 };
        b.Click += (_, _) => { using var d = new FolderBrowserDialog(); if (d.ShowDialog() == DialogResult.OK) tb.Text = d.SelectedPath; };
        parent.Controls.Add(b);
    }

    private void Log(string s) =>
        _log.BeginInvoke(() => { _log.AppendText(s + Environment.NewLine); });

    private void Busy(bool busy)
    {
        _btnExtract.Enabled = _btnRun.Enabled = !busy;
        UseWaitCursor = busy;
    }

    private async Task OnExtractAsync()
    {
        if (string.IsNullOrWhiteSpace(_olkPath.Text) || string.IsNullOrWhiteSpace(_workDir.Text))
        {
            MessageBox.Show("Bitte Olk-Profil und Arbeitsordner angeben."); return;
        }
        Busy(true);
        _cts = new CancellationTokenSource();
        try
        {
            var runner = new OlkExtractorRunner(Log)
            {
                DevScriptPath = Path.Combine(AppContext.BaseDirectory, "extractor", "olkextract.py"),
            };
            Log($"Extrahiere aus {_olkPath.Text} ...");
            var code = await runner.RunAsync(_olkPath.Text, _workDir.Text, _cts.Token);
            Log(code == 0 ? "Extraktion fertig." : $"Extraktion mit Code {code} beendet.");
        }
        catch (Exception ex) { Log("FEHLER: " + ex.Message); }
        finally { Busy(false); }
    }

    private async Task OnRunAsync()
    {
        var manifest = Path.Combine(_workDir.Text, "manifest.json");
        if (!File.Exists(manifest)) { MessageBox.Show("manifest.json nicht gefunden - erst Schritt 1 (Extrahieren)."); return; }

        Busy(true);
        _cts = new CancellationTokenSource();
        _progress.Value = 0;
        try
        {
            var src = new ExtractSource { OutDir = _workDir.Text };
            var opts = new RunOptions { DryRun = _dryRun.Checked, IncludeDeletedItems = _inclDeleted.Checked, IncludeDrafts = _inclDrafts.Checked };
            IMailTarget target = BuildTarget();
            var progress = new Progress<ExportProgress>(p =>
            {
                if (p.Total > 0) { _progress.Maximum = p.Total; _progress.Value = Math.Min(p.Done, p.Total); }
            });
            var exporter = new Exporter(src, opts, Log);
            await using (target)
                await exporter.RunAsync(target, progress, _cts.Token);
            Log("Übertragung abgeschlossen.");
        }
        catch (Exception ex) { Log("FEHLER: " + ex.Message); }
        finally { Busy(false); }
    }

    private IMailTarget BuildTarget() => _target.SelectedIndex switch
    {
        0 => new ImapTarget(new ImapOptions
        {
            Host = _imHost.Text.Trim(), Port = (int)_imPort.Value, UseSsl = _imSsl.Checked,
            User = _imUser.Text, Password = _imPass.Text,
            FolderPrefix = string.IsNullOrWhiteSpace(_imPrefix.Text) ? null : _imPrefix.Text.Trim(),
        }),
        1 => new MboxTarget(new MboxOptions { OutDir = string.IsNullOrWhiteSpace(_mbDir.Text) ? Path.Combine(_workDir.Text, "MBOX") : _mbDir.Text }),
        2 => new ExchangeEwsTarget(new EwsOptions { EwsUrl = _ewUrl.Text.Trim(), User = _ewUser.Text, Password = _ewPass.Text }),
        3 => new PstTarget(new PstOptions { PstPath = string.IsNullOrWhiteSpace(_pstPath.Text) ? Path.Combine(_workDir.Text, "olk-export.pst") : _pstPath.Text }),
        _ => throw new InvalidOperationException("Kein Ziel gewählt."),
    };
}
