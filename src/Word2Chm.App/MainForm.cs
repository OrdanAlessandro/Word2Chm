using System.Diagnostics;
using System.Text;
using Word2Chm.Core;
using Word2Chm.Core.Compilation;
using Word2Chm.Core.Generation;

namespace Word2Chm.App;

internal sealed class MainForm : Form
{
    private readonly TextBox _docxPath = new();
    private readonly TextBox _outputDirectory = new();
    private readonly TextBox _baseName = new();
    private readonly NumericUpDown _startContextId = new();
    private readonly NumericUpDown _pageLevel = new();
    private readonly TextBox _hhcPath = new();
    private readonly CheckBox _compileChm = new();
    private readonly CheckBox _openOutput = new();
    private readonly TextBox _log = new();
    private readonly Button _convertButton = new();
    private readonly ProgressBar _progress = new();

    private CancellationTokenSource? _cancellation;

    public MainForm()
    {
        Text = "Word2Chm - Da Word a guida HTML Help (.chm)";
        MinimumSize = new Size(760, 620);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9F);

        BuildLayout();
        LoadDefaults();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 9,
            Padding = new Padding(12),
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));

        var row = 0;

        // Input DOCX.
        root.Controls.Add(Label("Documento Word:"), 0, row);
        _docxPath.Dock = DockStyle.Fill;
        root.Controls.Add(_docxPath, 1, row);
        var browseDocx = Button("Sfoglia...", BrowseDocx);
        root.Controls.Add(browseDocx, 2, row++);

        // Output directory.
        root.Controls.Add(Label("Cartella di output:"), 0, row);
        _outputDirectory.Dock = DockStyle.Fill;
        root.Controls.Add(_outputDirectory, 1, row);
        var browseOutput = Button("Sfoglia...", BrowseOutput);
        root.Controls.Add(browseOutput, 2, row++);

        // Base name.
        root.Controls.Add(Label("Nome base:"), 0, row);
        _baseName.Dock = DockStyle.Fill;
        root.Controls.Add(_baseName, 1, row++);

        // Start context ID.
        root.Controls.Add(Label("ID di contesto iniziale:"), 0, row);
        _startContextId.Dock = DockStyle.Left;
        _startContextId.Minimum = 1;
        _startContextId.Maximum = 65535;
        _startContextId.Width = 120;
        root.Controls.Add(_startContextId, 1, row++);

        // Page level.
        root.Controls.Add(Label("Nuova pagina al livello:"), 0, row);
        _pageLevel.Dock = DockStyle.Left;
        _pageLevel.Minimum = 1;
        _pageLevel.Maximum = 6;
        _pageLevel.Width = 120;
        root.Controls.Add(_pageLevel, 1, row++);

        // hhc.exe.
        root.Controls.Add(Label("hhc.exe:"), 0, row);
        _hhcPath.Dock = DockStyle.Fill;
        root.Controls.Add(_hhcPath, 1, row);
        var browseHhc = Button("Rileva", DetectHhc);
        root.Controls.Add(browseHhc, 2, row++);

        // Options.
        var options = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        _compileChm.Text = "Compila il .chm";
        _compileChm.Checked = true;
        _compileChm.AutoSize = true;
        _openOutput.Text = "Apri la cartella al termine";
        _openOutput.AutoSize = true;
        options.Controls.Add(_compileChm);
        options.Controls.Add(_openOutput);
        root.Controls.Add(options, 1, row++);

        // Convert button.
        _convertButton.Text = "Converti";
        _convertButton.AutoSize = true;
        _convertButton.Click += async (_, _) => await RunConversionAsync();
        root.Controls.Add(_convertButton, 1, row++);

        // Log.
        root.Controls.Add(Label("Log:"), 0, row);
        _log.Multiline = true;
        _log.ReadOnly = true;
        _log.ScrollBars = ScrollBars.Both;
        _log.WordWrap = false;
        _log.Font = new Font("Consolas", 8.5F);
        _log.Dock = DockStyle.Fill;

        _progress.Dock = DockStyle.Bottom;
        _progress.Style = ProgressBarStyle.Marquee;
        _progress.Visible = false;

        var logPanel = new Panel { Dock = DockStyle.Fill };
        logPanel.Controls.Add(_log);
        logPanel.Controls.Add(_progress);
        root.Controls.Add(logPanel, 1, row);
        root.SetColumnSpan(logPanel, 2);

        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        Controls.Add(root);
        AcceptButton = _convertButton;
    }

    private static Label Label(string text) => new()
    {
        Text = text,
        AutoSize = true,
        Anchor = AnchorStyles.Left,
        TextAlign = ContentAlignment.MiddleLeft,
    };

    private static Button Button(string text, Action onClick)
    {
        var button = new Button { Text = text, AutoSize = true };
        button.Click += (_, _) => onClick();
        return button;
    }

    private void LoadDefaults()
    {
        var settings = AppSettings.Load();

        _docxPath.Text = settings.DocxPath ?? string.Empty;
        _outputDirectory.Text = settings.OutputDirectory ?? string.Empty;
        _startContextId.Value = Math.Clamp(settings.StartContextId, 1, 65535);
        _pageLevel.Value = Math.Clamp(settings.PageLevel, 1, 6);
        _hhcPath.Text = settings.HhcPath ?? HhcLocator.Locate() ?? string.Empty;

        if (string.IsNullOrEmpty(_hhcPath.Text))
        {
            AppendLog("hhc.exe non rilevato automaticamente: usa il pulsante \"Rileva\" o specifica il percorso.");
        }
    }

    private void SaveSettings() => new AppSettings
    {
        DocxPath = _docxPath.Text,
        OutputDirectory = _outputDirectory.Text,
        StartContextId = (int)_startContextId.Value,
        PageLevel = (int)_pageLevel.Value,
        HhcPath = _hhcPath.Text,
    }.Save();

    private void BrowseDocx()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Documenti Word (*.docx)|*.docx|Tutti i file (*.*)|*.*",
            Title = "Seleziona il documento Word",
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _docxPath.Text = dialog.FileName;
            if (string.IsNullOrWhiteSpace(_baseName.Text))
            {
                _baseName.Text = Path.GetFileNameWithoutExtension(dialog.FileName);
            }

            if (string.IsNullOrWhiteSpace(_outputDirectory.Text))
            {
                _outputDirectory.Text = Path.Combine(
                    Path.GetDirectoryName(dialog.FileName) ?? Environment.CurrentDirectory,
                    Path.GetFileNameWithoutExtension(dialog.FileName) + "-chm");
            }
        }
    }

    private void BrowseOutput()
    {
        using var dialog = new FolderBrowserDialog { Description = "Seleziona la cartella di output" };
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _outputDirectory.Text = dialog.SelectedPath;
        }
    }

    private void DetectHhc()
    {
        var found = HhcLocator.Locate(_hhcPath.Text);
        if (found is not null)
        {
            _hhcPath.Text = found;
            AppendLog("hhc.exe trovato: " + found);
        }
        else
        {
            AppendLog("hhc.exe non trovato. Installa HTML Help Workshop oppure indica il percorso manualmente.");
        }
    }

    private async Task RunConversionAsync()
    {
        if (ValidateInputs() is { } error)
        {
            MessageBox.Show(this, error, "Dati mancanti", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SaveSettings();
        SetBusy(true);

        var options = new ConversionOptions
        {
            DocxPath = _docxPath.Text.Trim(),
            OutputDirectory = _outputDirectory.Text.Trim(),
            BaseName = string.IsNullOrWhiteSpace(_baseName.Text) ? null : _baseName.Text.Trim(),
            Build = new BuildOptions
            {
                DefaultContextId = (int)_startContextId.Value,
                PageLevel = (int)_pageLevel.Value,
            },
            Compile = new CompileOptions
            {
                HhcPath = _compileChm.Checked ? _hhcPath.Text.Trim() : null,
            },
        };

        _cancellation = new CancellationTokenSource();
        var token = _cancellation.Token;

        try
        {
            var result = await Task.Run(() => new ConversionPipeline().Run(options), token);
            ReportResult(result);

            if (_openOutput.Checked)
            {
                OpenInExplorer(result.OutputDirectory);
            }
        }
        catch (Exception ex)
        {
            AppendLog("ERRORE: " + ex.Message);
            MessageBox.Show(this, ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _cancellation.Dispose();
            _cancellation = null;
            SetBusy(false);
        }
    }

    private void ReportResult(ConversionResult result)
    {
        AppendLog($"Documento: {result.Document.Title}");
        AppendLog($"Pagine generate: {result.Document.Pages.Count}");

        var withIds = result.Document.Pages.Count(p => p.Symbol is not null);
        AppendLog($"ID di contesto definiti: {withIds}");
        AppendLog($"Voci di indice: {result.Document.IndexEntries.Count}");
        AppendLog("File generati:");
        foreach (var file in result.GeneratedFiles)
        {
            AppendLog("  " + Path.GetRelativePath(result.OutputDirectory, file));
        }

        if (result.Compilation.Success)
        {
            AppendLog("Compilazione CHM completata: " + result.ChmPath);
        }
        else
        {
            AppendLog("Compilazione CHM non eseguita o non riuscita.");
            if (!string.IsNullOrWhiteSpace(result.Compilation.Output))
            {
                AppendLog(result.Compilation.Output);
            }
        }
    }

    private string? ValidateInputs()
    {
        if (string.IsNullOrWhiteSpace(_docxPath.Text) || !File.Exists(_docxPath.Text))
        {
            return "Seleziona un documento Word esistente.";
        }

        if (string.IsNullOrWhiteSpace(_outputDirectory.Text))
        {
            return "Seleziona la cartella di output.";
        }

        if (_compileChm.Checked && !string.IsNullOrWhiteSpace(_hhcPath.Text) && !File.Exists(_hhcPath.Text))
        {
            return "Il percorso di hhc.exe non esiste. Correggilo oppure disabilita la compilazione.";
        }

        return null;
    }

    private void SetBusy(bool busy)
    {
        _convertButton.Enabled = !busy;
        _progress.Visible = busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
    }

    private void AppendLog(string message)
    {
        if (_log.InvokeRequired)
        {
            _log.BeginInvoke(() => AppendLog(message));
            return;
        }

        _log.AppendText(message + Environment.NewLine);
    }

    private void OpenInExplorer(string directory)
    {
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{directory}\"") { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            AppendLog("Impossibile aprire la cartella: " + ex.Message);
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _cancellation?.Cancel();
        base.OnFormClosing(e);
    }
}

/// <summary>Small JSON-backed user settings store kept next to the executable.</summary>
internal sealed class AppSettings
{
    private static string FilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Word2Chm",
        "settings.json");

    public string? DocxPath { get; set; }
    public string? OutputDirectory { get; set; }
    public string? HhcPath { get; set; }
    public int StartContextId { get; set; } = 1000;
    public int PageLevel { get; set; } = 1;

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                return System.Text.Json.JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(FilePath))
                       ?? new AppSettings();
            }
        }
        catch (Exception)
        {
            // Corrupt settings should never block startup.
        }

        return new AppSettings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(
                FilePath,
                System.Text.Json.JsonSerializer.Serialize(this, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true,
                }));
        }
        catch (Exception)
        {
            // Saving settings is best-effort.
        }
    }
}
