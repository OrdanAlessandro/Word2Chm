using System.Diagnostics;
using System.Text;
using Word2Chm.Core;
using Word2Chm.Core.Common;
using Word2Chm.Core.Compilation;
using Word2Chm.Core.Generation;
using Word2Chm.Core.Model;

namespace Word2Chm.App;

internal sealed class MainForm : Form
{
    private readonly TextBox _docxPath = new();
    private readonly TextBox _outputDirectory = new();
    private readonly TextBox _baseName = new();
    private readonly TextBox _chmFileName = new();
    private readonly TextBox _contextIdHeaderPath = new();
    private readonly NumericUpDown _pageLevel = new();
    private readonly NumericUpDown _bodyFontSize = new();
    private readonly TextBox _templateDirectory = new();
    private readonly TextBox _footer = new();
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
            RowCount = 12,
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

        // Header with the ID definitions. The document names the symbols, the numbers come
        // from this file, so the same IDs hold across every language edition.
        root.Controls.Add(Label("File ID (.h):"), 0, row);
        _contextIdHeaderPath.Dock = DockStyle.Fill;
        root.Controls.Add(_contextIdHeaderPath, 1, row);
        var browseHeader = Button("Sfoglia...", BrowseContextIdHeader);
        root.Controls.Add(browseHeader, 2, row++);

        // Name of the compiled .chm.
        root.Controls.Add(Label("File CHM:"), 0, row);
        _chmFileName.Dock = DockStyle.Fill;
        root.Controls.Add(_chmFileName, 1, row);
        root.SetColumnSpan(_chmFileName, 2);
        row++;

        // Page level.
        root.Controls.Add(Label("Nuova pagina al livello:"), 0, row);
        _pageLevel.Dock = DockStyle.Left;
        _pageLevel.Minimum = 1;
        _pageLevel.Maximum = 6;
        _pageLevel.Width = 120;
        root.Controls.Add(_pageLevel, 1, row++);

        // Body text size.
        root.Controls.Add(Label("Dimensione testo (pt):"), 0, row);
        _bodyFontSize.Dock = DockStyle.Left;
        _bodyFontSize.Minimum = 6;
        _bodyFontSize.Maximum = 24;
        _bodyFontSize.DecimalPlaces = 1;
        _bodyFontSize.Increment = 0.5m;
        _bodyFontSize.Width = 120;
        root.Controls.Add(_bodyFontSize, 1, row++);

        // WinCHM skin.
        root.Controls.Add(Label("Template (skin):"), 0, row);
        _templateDirectory.Dock = DockStyle.Fill;
        root.Controls.Add(_templateDirectory, 1, row);
        var browseTemplate = Button("Sfoglia...", BrowseTemplate);
        root.Controls.Add(browseTemplate, 2, row++);

        // Footer.
        root.Controls.Add(Label("Piè di pagina:"), 0, row);
        _footer.Dock = DockStyle.Fill;
        root.Controls.Add(_footer, 1, row);
        root.SetColumnSpan(_footer, 2);
        row++;

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
        _pageLevel.Value = Math.Clamp(settings.PageLevel, 1, 6);
        _bodyFontSize.Value = Math.Clamp(
            (decimal)settings.BodyFontSizePt,
            _bodyFontSize.Minimum,
            _bodyFontSize.Maximum);
        _templateDirectory.Text = settings.TemplateDirectory ?? DefaultTemplateDirectory() ?? string.Empty;
        _contextIdHeaderPath.Text = settings.ContextIdHeaderPath ?? string.Empty;
        _chmFileName.Text = ChmProjectNames.NormalizeChmFileName(settings.ChmFileName);
        _footer.Text = settings.Footer ?? HelpDocument.DefaultFooter;
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
        PageLevel = (int)_pageLevel.Value,
        BodyFontSizePt = (double)_bodyFontSize.Value,
        TemplateDirectory = string.IsNullOrWhiteSpace(_templateDirectory.Text) ? null : _templateDirectory.Text.Trim(),
        ContextIdHeaderPath = string.IsNullOrWhiteSpace(_contextIdHeaderPath.Text) ? null : _contextIdHeaderPath.Text.Trim(),
        ChmFileName = ChmProjectNames.NormalizeChmFileName(_chmFileName.Text),
        Footer = _footer.Text,
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

    private void BrowseContextIdHeader()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "File header (*.h)|*.h|Tutti i file (*.*)|*.*",
            Title = "Seleziona il file .h con le definizioni degli ID",
            FileName = _contextIdHeaderPath.Text,
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _contextIdHeaderPath.Text = dialog.FileName;
        }
    }

    private void BrowseTemplate()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Seleziona la cartella del template (contiene fixedtop.htm)",
            SelectedPath = Directory.Exists(_templateDirectory.Text) ? _templateDirectory.Text : string.Empty,
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _templateDirectory.Text = dialog.SelectedPath;
        }
    }

    /// <summary>
    /// Looks for the bundled skin next to the executable, then in the repository layout used
    /// during development, so the template works without a manual selection.
    /// </summary>
    private static string? DefaultTemplateDirectory()
    {
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "template", "fixedtop"),
            Path.Combine(AppContext.BaseDirectory, "template"),
        };

        foreach (var candidate in candidates)
        {
            if (File.Exists(Path.Combine(candidate, WinChmTemplate.TemplateFileName)))
            {
                return candidate;
            }
        }

        return null;
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
            ContextIdHeaderPath = string.IsNullOrWhiteSpace(_contextIdHeaderPath.Text) ? null : _contextIdHeaderPath.Text.Trim(),
            Build = new BuildOptions
            {
                PageLevel = (int)_pageLevel.Value,
                BodyFontSizePt = (double)_bodyFontSize.Value,
                Footer = string.IsNullOrWhiteSpace(_footer.Text) ? HelpDocument.DefaultFooter : _footer.Text.Trim(),
            },
            TemplateDirectory = string.IsNullOrWhiteSpace(_templateDirectory.Text) ? null : _templateDirectory.Text.Trim(),
            ChmFileName = _chmFileName.Text,
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
        catch (DocxUnreadableException ex)
        {
            // A missing, locked or non-Word document is the user's mistake to fix, so it gets
            // a clear popup instead of the generic error path below.
            AppendLog("ERRORE: " + ex.Message);
            MessageBox.Show(this, ex.Message, "Documento non leggibile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        var anchoredIds = result.Document.Pages.Sum(p => p.Anchors.Count);
        var resolved = result.Document.Pages.Count(p => p.ContextId.HasValue) +
                       result.Document.Pages.Sum(p => p.Anchors.Count(a => a.ContextId.HasValue));
        AppendLog($"ID di contesto definiti: {withIds + anchoredIds} (risolti dal file .h: {resolved})");
        if (anchoredIds > 0)
        {
            AppendLog($"  di cui su sottotitoli (collegati a un'ancora): {anchoredIds}");
        }

        AppendLog($"Voci di indice: {result.Document.IndexEntries.Count}");

        foreach (var warning in result.Document.Warnings)
        {
            AppendLog("AVVISO: " + warning);
        }

        if (result.MissingIdsReportPath is not null)
        {
            AppendLog("ATTENZIONE: alcuni ID usati nel documento non sono definiti nel file .h.");
            AppendLog("Elenco scritto in: " + Path.GetRelativePath(result.OutputDirectory, result.MissingIdsReportPath));
        }

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

        // The missing-ID list is what the user has to act on, so open it right away instead
        // of making them hunt for it in the output folder.
        if (result.MissingIdsReportPath is not null)
        {
            OpenWithDefaultEditor(result.MissingIdsReportPath);
        }
    }

    /// <summary>Opens a text file in the default editor, falling back to Notepad.</summary>
    private static void OpenWithDefaultEditor(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception)
        {
            // open with notepad as fallback
            Process.Start("notepad.exe", path);
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

        if (string.IsNullOrWhiteSpace(_contextIdHeaderPath.Text) || !File.Exists(_contextIdHeaderPath.Text))
        {
            return "Seleziona il file .h con le definizioni degli ID di contesto.";
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
    /// <summary>
    /// Version of the stored layout. Files written before <see cref="CurrentVersion"/> may
    /// carry a <see cref="PageLevel"/> that only reflects the old default, so it is discarded
    /// instead of silently pinning every conversion back to one big page.
    /// </summary>
    private const int CurrentVersion = 2;

    private static string FilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Word2Chm",
        "settings.json");

    public int Version { get; set; }

    public string? DocxPath { get; set; }
    public string? OutputDirectory { get; set; }
    public string? HhcPath { get; set; }
    public int PageLevel { get; set; } = BuildOptions.DefaultPageLevel;

    /// <summary>
    /// Body text size in points. A file written before this setting existed has no such
    /// property, so the initializer supplies the default; a stored value outside the
    /// allowed range is discarded on load.
    /// </summary>
    public double BodyFontSizePt { get; set; } = BuildOptions.DefaultBodyFontSizePt;
    public string? TemplateDirectory { get; set; }

    /// <summary>
    /// Path of the C++ header (e.g. helpId.h) the numeric context IDs are read from.
    /// </summary>
    public string? ContextIdHeaderPath { get; set; }

    /// <summary>
    /// Name of the compiled .chm. Null for files written before this setting existed,
    /// where the default name still applies.
    /// </summary>
    public string? ChmFileName { get; set; }

    /// <summary>
    /// Footer written to every topic. A file from before this setting existed has a null
    /// value, so the default is applied on load rather than leaving a blank footer.
    /// </summary>
    public string? Footer { get; set; } = HelpDocument.DefaultFooter;

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                var settings = System.Text.Json.JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(FilePath));
                if (settings is not null)
                {
                    if (settings.Version < CurrentVersion)
                    {
                        settings.PageLevel = BuildOptions.DefaultPageLevel;
                    }

                    if (settings.BodyFontSizePt is < 6 or > 24)
                    {
                        settings.BodyFontSizePt = BuildOptions.DefaultBodyFontSizePt;
                    }

                    return settings;
                }
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
        Version = CurrentVersion;
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
