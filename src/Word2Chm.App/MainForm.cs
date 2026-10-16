using System.Diagnostics;
using System.Text;
using Word2Chm.Core;
using Word2Chm.Core.Common;
using Word2Chm.Core.Compilation;
using Word2Chm.Core.Generation;
using Word2Chm.Core.Model;
using Word2Chm.Core.Project;

namespace Word2Chm.App;

internal sealed partial class MainForm : Form
{
    private CancellationTokenSource? _cancellation;

    /// <summary>
    /// UI state that is not a conversion input: which .w2c is open and whether the fields
    /// have changed since it was last saved. Both change the title bar, which is where the
    /// user reads them.
    /// </summary>
    private string? _projectPath;
    private string _savedProjectJson = string.Empty;

    private readonly string? _startupProjectPath;

    // Visual Studio instantiates this constructor without reading settings or opening files.
    public MainForm()
    {
        InitializeComponent();
    }

    public MainForm(string? projectPath) : this()
    {
        _startupProjectPath = projectPath;
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        if (DesignMode || System.ComponentModel.LicenseManager.UsageMode ==
            System.ComponentModel.LicenseUsageMode.Designtime)
        {
            return;
        }

        ApplyWindowIcon();
        LoadToolbarIcons();
        LoadDefaults();
        _savedProjectJson = CurrentProject().Serialize();
        if (!string.IsNullOrWhiteSpace(_startupProjectPath))
        {
            OpenProjectAtStartup(_startupProjectPath);
        }
        UpdateTitle();
    }

    /// <summary>
    /// Opens the project passed on the command line, i.e. the file that was double-clicked.
    /// A project that cannot be read is reported and the window stays on its defaults rather
    /// than closing, so the user still has a working application.
    /// </summary>
    private void OpenProjectAtStartup(string path)
    {
        try
        {
            ApplyProject(ProjectFile.Load(path));
            AppendLog("Progetto aperto: " + Path.GetFullPath(path));
        }
        catch (ProjectFileException ex)
        {
            AppendLog("ERRORE: " + ex.Message);
            MessageBox.Show(this, ex.Message, "Progetto non leggibile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    /// <summary>
    /// Uses the executable icon for the title bar and task bar. Loaded from the file next to
    /// the executable rather than the embedded resource so the same asset drives the window,
    /// the .exe and the .w2c association.
    /// </summary>
    private void ApplyWindowIcon()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "app.ico");
        try
        {
            if (File.Exists(path))
            {
                Icon = new Icon(path);
            }
        }
        catch (Exception)
        {
            // A missing or unreadable icon must not stop the application from starting.
        }
    }

    /// <summary>
    /// New/open/save, placed above the fields so the file commands sit where the title bar
    /// suggests. Icons are 16px PNGs shipped next to the executable; a missing file only
    /// costs the picture, since every command keeps its label.
    /// </summary>
    private void LoadToolbarIcons()
    {
        // The Designer may omit the container when there are no design-time components.
        // Runtime-created icons still need an owner that disposes them with the form.
        components ??= new System.ComponentModel.Container();
        var icons = BuildIconList();
        components.Add(icons);
        _toolStrip.ImageList = icons;
        foreach (ToolStripItem item in _toolStrip.Items)
        {
            item.DisplayStyle = icons.Images.ContainsKey(item.ImageKey)
                ? ToolStripItemDisplayStyle.ImageAndText
                : ToolStripItemDisplayStyle.Text;
        }
    }

    private static ImageList BuildIconList()
    {
        var list = new ImageList
        {
            ImageSize = new Size(16, 16),
            ColorDepth = ColorDepth.Depth32Bit,
            TransparentColor = Color.Transparent,
        };

        foreach (var name in new[] { "new", "open", "save", "save-as" })
        {
            var path = Path.Combine(AppContext.BaseDirectory, "icons", $"{name}-16.png");
            if (!File.Exists(path))
            {
                continue;
            }

            try
            {
                using var stream = File.OpenRead(path);
                // Copied into a new bitmap so the list does not hold the file open.
                list.Images.Add(name, new Bitmap(Image.FromStream(stream)));
            }
            catch (Exception)
            {
                // Skip the picture; the text label still identifies the command.
            }
        }

        return list;
    }

    private void NewProject_Click(object? sender, EventArgs e) => NewProject();
    private void OpenProject_Click(object? sender, EventArgs e) => OpenProject();
    private void SaveProject_Click(object? sender, EventArgs e) => SaveProject(saveAs: false);
    private void SaveProjectAs_Click(object? sender, EventArgs e) => SaveProject(saveAs: true);
    private void BrowseDocx_Click(object? sender, EventArgs e) => BrowseDocx();
    private void BrowseOutput_Click(object? sender, EventArgs e) => BrowseOutput();
    private void BrowseHeader_Click(object? sender, EventArgs e) => BrowseContextIdHeader();
    private void BrowseChmCopy_Click(object? sender, EventArgs e) => BrowseChmCopyPath();
    private void BrowseTemplate_Click(object? sender, EventArgs e) => BrowseTemplate();
    private void DetectHhc_Click(object? sender, EventArgs e) => DetectHhc();
    private async void Convert_Click(object? sender, EventArgs e) => await RunConversionAsync();

    private void LoadDefaults()
    {
        var settings = AppSettings.Load();

        _docxPath.Text = settings.DocxPath ?? string.Empty;
        _outputDirectory.Text = settings.OutputDirectory ?? string.Empty;
        _pageLevel.Value = Math.Clamp(settings.PageLevel, BuildOptions.MinPageLevel, BuildOptions.MaxPageLevel);
        _bodyFontSize.Value = Math.Clamp(
            (decimal)settings.BodyFontSizePt,
            _bodyFontSize.Minimum,
            _bodyFontSize.Maximum);
        _templateDirectory.Text = settings.TemplateDirectory ?? DefaultTemplateDirectory() ?? string.Empty;
        _contextIdHeaderPath.Text = settings.ContextIdHeaderPath ?? string.Empty;
        _chmFileName.Text = ChmProjectNames.NormalizeChmFileName(settings.ChmFileName);
        _chmCopyPath.Text = settings.ChmCopyPath ?? string.Empty;
        _footer.Text = settings.Footer ?? HelpDocument.DefaultFooter;
        _hhcPath.Text = settings.HhcPath ?? HhcLocator.Locate() ?? string.Empty;
        _openChm.Checked = settings.OpenChmOnFinish;

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
        TemplateDirectory = Blank(_templateDirectory.Text),
        ContextIdHeaderPath = Blank(_contextIdHeaderPath.Text),
        ChmFileName = ChmProjectNames.NormalizeChmFileName(_chmFileName.Text),
        ChmCopyPath = Blank(_chmCopyPath.Text),
        Footer = _footer.Text,
        HhcPath = _hhcPath.Text,
        OpenChmOnFinish = _openChm.Checked,
    }.Save();

    private static string? Blank(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    /// <summary>Fills the fields from a project, replacing whatever is there.</summary>
    private void ApplyProject(ProjectFile project)
    {
        _docxPath.Text = project.DocxPath ?? string.Empty;
        _outputDirectory.Text = project.OutputDirectory ?? string.Empty;
        _baseName.Text = project.BaseName ?? string.Empty;
        _contextIdHeaderPath.Text = project.ContextIdHeaderPath ?? string.Empty;
        _chmFileName.Text = ChmProjectNames.NormalizeChmFileName(project.ChmFileName);
        _chmCopyPath.Text = project.ChmCopyPath ?? string.Empty;
        _pageLevel.Value = Math.Clamp(project.PageLevel, BuildOptions.MinPageLevel, BuildOptions.MaxPageLevel);
        _bodyFontSize.Value = Math.Clamp(
            (decimal)project.BodyFontSizePt,
            _bodyFontSize.Minimum,
            _bodyFontSize.Maximum);
        _templateDirectory.Text = project.TemplateDirectory ?? string.Empty;
        _footer.Text = project.Footer ?? HelpDocument.DefaultFooter;
        _hhcPath.Text = project.HhcPath ?? string.Empty;
        _compileChm.Checked = project.CompileChm;
        _openOutput.Checked = project.OpenOutputOnFinish;
        _openChm.Checked = project.OpenChmOnFinish;

        _projectPath = project.SourcePath;
        _savedProjectJson = CurrentProject().Serialize();
        UpdateTitle();
    }

    /// <summary>Snapshots the fields as a project, without touching the open file.</summary>
    private ProjectFile CurrentProject() => new()
    {
        SourcePath = _projectPath,
        DocxPath = Blank(_docxPath.Text),
        OutputDirectory = Blank(_outputDirectory.Text),
        BaseName = Blank(_baseName.Text),
        ContextIdHeaderPath = Blank(_contextIdHeaderPath.Text),
        ChmFileName = _chmFileName.Text,
        ChmCopyPath = Blank(_chmCopyPath.Text),
        PageLevel = (int)_pageLevel.Value,
        BodyFontSizePt = (double)_bodyFontSize.Value,
        TemplateDirectory = Blank(_templateDirectory.Text),
        Footer = _footer.Text,
        HhcPath = Blank(_hhcPath.Text),
        CompileChm = _compileChm.Checked,
        OpenOutputOnFinish = _openOutput.Checked,
        OpenChmOnFinish = _openChm.Checked,
    };

    /// <summary>
    /// True when the fields differ from the project last saved or opened. Compared through
    /// the serialized form so a new field cannot silently go unnoticed.
    /// </summary>
    private bool HasUnsavedChanges() =>
        !string.Equals(CurrentProject().Serialize(), _savedProjectJson, StringComparison.Ordinal);

    private void NewProject()
    {
        if (!ConfirmDiscardChanges())
        {
            return;
        }

        SaveSettings();

        // A new project starts from the same defaults as a first run, so the fields come back
        // to a known state rather than keeping the previous project's values.
        _docxPath.Clear();
        _outputDirectory.Clear();
        _baseName.Clear();
        _contextIdHeaderPath.Clear();
        _chmCopyPath.Clear();
        _pageLevel.Value = BuildOptions.DefaultPageLevel;
        _bodyFontSize.Value = (decimal)BuildOptions.DefaultBodyFontSizePt;
        _footer.Text = HelpDocument.DefaultFooter;
        _templateDirectory.Text = DefaultTemplateDirectory() ?? string.Empty;
        _chmFileName.Text = ChmProjectNames.NormalizeChmFileName(null);
        _hhcPath.Text = HhcLocator.Locate() ?? string.Empty;
        _compileChm.Checked = true;
        _openOutput.Checked = false;
        _openChm.Checked = false;

        _projectPath = null;
        _savedProjectJson = CurrentProject().Serialize();
        UpdateTitle();
        AppendLog("Nuovo progetto.");
    }

    private void OpenProject()
    {
        if (!ConfirmDiscardChanges())
        {
            return;
        }

        using var dialog = new OpenFileDialog
        {
            Filter = ProjectFile.DialogFilter,
            Title = "Apri un progetto Word2Chm",
            DefaultExt = ProjectFile.Extension,
            CheckFileExists = true,
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            ApplyProject(ProjectFile.Load(dialog.FileName));
            AppendLog("Progetto aperto: " + dialog.FileName);
        }
        catch (ProjectFileException ex)
        {
            AppendLog("ERRORE: " + ex.Message);
            MessageBox.Show(this, ex.Message, "Progetto non leggibile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void SaveProject(bool saveAs)
    {
        var target = _projectPath;

        if (saveAs || string.IsNullOrWhiteSpace(target))
        {
            using var dialog = new SaveFileDialog
            {
                Filter = ProjectFile.DialogFilter,
                Title = saveAs ? "Salva il progetto con un altro nome" : "Salva il progetto",
                DefaultExt = ProjectFile.Extension,
                FileName = string.IsNullOrWhiteSpace(target)
                    ? ProjectFile.SuggestedFileName(_docxPath.Text)
                    : Path.GetFileName(target),
                InitialDirectory = string.IsNullOrWhiteSpace(target)
                    ? SafeDirectory(_docxPath.Text)
                    : Path.GetDirectoryName(target),
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            target = dialog.FileName;
        }

        var project = CurrentProject();
        try
        {
            project.Save(target!);
        }
        catch (ProjectFileException ex)
        {
            AppendLog("ERRORE: " + ex.Message);
            MessageBox.Show(this, ex.Message, "Salvataggio non riuscito", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _projectPath = project.SourcePath;
        _savedProjectJson = CurrentProject().Serialize();
        UpdateTitle();
        AppendLog("Progetto salvato: " + _projectPath);
    }

    /// <summary>Folder of a path, for seeding a dialog; an empty string when there is none.</summary>
    private static string SafeDirectory(string path)
    {
        try
        {
            return string.IsNullOrWhiteSpace(path) ? string.Empty : Path.GetDirectoryName(path) ?? string.Empty;
        }
        catch (ArgumentException)
        {
            // A half-typed path is not worth an error when it only seeds a dialog.
            return string.Empty;
        }
    }

    /// <summary>Uses the current path, or its nearest existing parent, to open a picker.</summary>
    private static string ExistingDialogDirectory(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        try
        {
            var directory = Path.GetFullPath(path.Trim());
            while (!Directory.Exists(directory))
            {
                directory = Path.GetDirectoryName(directory);
                if (string.IsNullOrEmpty(directory))
                {
                    return string.Empty;
                }
            }
            return directory;
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or IOException)
        {
            return string.Empty;
        }
    }

    private static void InitializeFileDialog(FileDialog dialog, string currentPath)
    {
        dialog.RestoreDirectory = true;
        dialog.InitialDirectory = ExistingDialogDirectory(currentPath);
        if (string.IsNullOrWhiteSpace(currentPath) || Directory.Exists(currentPath.Trim()))
        {
            return;
        }

        try
        {
            var name = Path.GetFileName(currentPath.Trim());
            if (!string.IsNullOrEmpty(name) && name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0)
            {
                dialog.FileName = name;
            }
        }
        catch (ArgumentException)
        {
            // Ignore an invalid, partially typed path when initializing the picker.
        }
    }

    /// <summary>
    /// Asks before throwing away edits. Returns true when it is safe to continue.
    /// </summary>
    private bool ConfirmDiscardChanges()
    {
        if (!HasUnsavedChanges())
        {
            return true;
        }

        var answer = MessageBox.Show(
            this,
            "Il progetto ha modifiche non salvate. Salvarle prima di continuare?",
            "Modifiche non salvate",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        switch (answer)
        {
            case DialogResult.Yes:
                SaveProject(saveAs: false);
                // If the save was cancelled the fields are still dirty, so continue only when
                // it actually went through.
                return !HasUnsavedChanges();
            case DialogResult.No:
                return true;
            default:
                return false;
        }
    }

    private void UpdateTitle()
    {
        var name = string.IsNullOrWhiteSpace(_projectPath)
            ? "Senza titolo"
            : Path.GetFileName(_projectPath);
        var dirty = HasUnsavedChanges() ? " *" : string.Empty;
        Text = $"Word2Chm - {name}{dirty}";
    }

    private void BrowseDocx()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Documenti Word (*.docx)|*.docx|Tutti i file (*.*)|*.*",
            Title = "Seleziona il documento Word",
        };
        InitializeFileDialog(dialog, _docxPath.Text);

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
        using var dialog = new FolderBrowserDialog
        {
            Description = "Seleziona la cartella di output",
            SelectedPath = ExistingDialogDirectory(_outputDirectory.Text),
        };
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
        };
        InitializeFileDialog(dialog, _contextIdHeaderPath.Text);

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _contextIdHeaderPath.Text = dialog.FileName;
        }
    }

    /// <summary>
    /// Picks the destination for the .chm copy. A file dialog is used by default because the
    /// common case is dropping the guide into a folder the application already defines; typing a
    /// directory into the field still works and is honoured by the pipeline.
    /// </summary>
    private void BrowseChmCopyPath()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Guida compilata (*.chm)|*.chm|Tutti i file (*.*)|*.*",
            Title = "Dove copiare il file .chm al termine",
            FileName = ChmProjectNames.NormalizeChmFileName(_chmFileName.Text),
            OverwritePrompt = false,
        };
        InitializeFileDialog(dialog, _chmCopyPath.Text);

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _chmCopyPath.Text = dialog.FileName;
        }
    }

    private void BrowseTemplate()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "Seleziona la cartella del template (contiene fixedtop.htm)",
            SelectedPath = ExistingDialogDirectory(_templateDirectory.Text),
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

        // The fields are the project, so the conversion input is derived from the same
        // snapshot the file commands use; there is no second place where they could drift.
        var project = CurrentProject();
        var options = project.ToConversionOptions();

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
            if (project.OpenChmOnFinish && result.Compilation.Success && result.ChmPath is not null)
            {
                OpenChm(result.ChmPath);
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

    private void OpenChm(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or
                                   InvalidOperationException or IOException or UnauthorizedAccessException)
        {
            AppendLog("AVVISO: il CHM è stato compilato ma non è stato possibile aprirlo: " + ex.Message);
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
            if (result.CopiedChmPath is not null)
            {
                AppendLog("Copia del .chm scritta in: " + result.CopiedChmPath);
            }
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

        // There is nothing to copy when the .chm is not compiled, so catch the combination
        // here rather than silently leaving the requested copy undone.
        if (!_compileChm.Checked && !string.IsNullOrWhiteSpace(_chmCopyPath.Text))
        {
            return "Per copiare il .chm devi anche abilitare \"Compila il .chm\".";
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
        // Closing with unsaved edits asks first, and a cancelled close leaves the window open.
        if (e.CloseReason == CloseReason.UserClosing && !ConfirmDiscardChanges())
        {
            e.Cancel = true;
            return;
        }

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
    public bool OpenChmOnFinish { get; set; }
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
    /// Destination the compiled .chm is copied to at the end of a successful conversion.
    /// Null means no copy is made.
    /// </summary>
    public string? ChmCopyPath { get; set; }

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
