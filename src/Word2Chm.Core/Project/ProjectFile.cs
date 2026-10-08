using System.Text.Json;
using System.Text.Json.Serialization;
using Word2Chm.Core.Compilation;
using Word2Chm.Core.Generation;
using Word2Chm.Core.Model;

namespace Word2Chm.Core.Project;

/// <summary>Raised when a .w2c project file cannot be read as a project.</summary>
public sealed class ProjectFileException : Exception
{
    public ProjectFileException(string message, Exception? inner = null) : base(message, inner)
    {
    }
}

/// <summary>
/// The inputs of one conversion, saved as a .w2c file so a setup can be reopened later or
/// handed to someone else. A .w2c holds paths, not the document, and lives next to the
/// document it describes in the usual case, so the two paths that point there — the Word
/// document and the .h header — are stored relative to the project file and resolved again
/// on load. That keeps a project folder usable after being moved or copied to another
/// machine, as long as the files travel with it.
/// </summary>
public sealed class ProjectFile
{
    /// <summary>Extension of a saved project, written without the leading dot.</summary>
    public const string Extension = "w2c";

    /// <summary>Filter for the open/save dialogs, e.g. <c>Progetti Word2Chm (*.w2c)|*.w2c</c>.</summary>
    public const string DialogFilter = "Progetti Word2Chm (*.w2c)|*.w2c|Tutti i file (*.*)|*.*";

    /// <summary>
    /// Version of the layout. Bumped when a stored value changes meaning; see
    /// <see cref="Load"/> for what an older file means.
    /// </summary>
    public const int CurrentVersion = 1;

    [JsonPropertyName("version")]
    public int Version { get; set; } = CurrentVersion;

    public string? DocxPath { get; set; }
    public string? OutputDirectory { get; set; }
    public string? BaseName { get; set; }
    public string? ContextIdHeaderPath { get; set; }
    public string? ChmFileName { get; set; }
    public string? ChmCopyPath { get; set; }
    public int PageLevel { get; set; } = BuildOptions.DefaultPageLevel;
    public double BodyFontSizePt { get; set; } = BuildOptions.DefaultBodyFontSizePt;
    public string? Footer { get; set; } = HelpDocument.DefaultFooter;
    public string? TemplateDirectory { get; set; }
    public string? HhcPath { get; set; }

    /// <summary>
    /// Whether the conversion should run the compiler. Kept apart from <see cref="HhcPath"/>
    /// so the discovered path survives a phase in which compilation is switched off.
    /// </summary>
    public bool CompileChm { get; set; } = true;

    /// <summary>Whether the output folder is opened when the conversion finishes.</summary>
    public bool OpenOutputOnFinish { get; set; }

    private static readonly JsonSerializerOptions WriteOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private static readonly JsonSerializerOptions ReadOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        // A .w2c is meant to be editable by hand, so a difference in casing is not an error.
        PropertyNameCaseInsensitive = true,
    };

    /// <summary>Full path of the project file this instance was loaded from, if any.</summary>
    [JsonIgnore]
    public string? SourcePath { get; set; }

    /// <summary>Builds the pipeline input described by this project.</summary>
    public ConversionOptions ToConversionOptions()
    {
        // Clamp here rather than trusting the file: a hand-edited .w2c is a normal way to
        // reuse a project, and a level outside the pages the viewer understands would
        // otherwise reach the builder.
        var pageLevel = Math.Clamp(PageLevel, 1, BuildOptions.MaxPageLevel);
        var bodyFontSize = BodyFontSizePt is >= BuildOptions.MinBodyFontSizePt and <= BuildOptions.MaxBodyFontSizePt
            ? BodyFontSizePt
            : BuildOptions.DefaultBodyFontSizePt;

        return new ConversionOptions
        {
            DocxPath = DocxPath ?? string.Empty,
            OutputDirectory = OutputDirectory ?? string.Empty,
            BaseName = Blank(BaseName),
            ContextIdHeaderPath = Blank(ContextIdHeaderPath),
            ChmFileName = Blank(ChmFileName),
            ChmCopyPath = Blank(ChmCopyPath),
            TemplateDirectory = Blank(TemplateDirectory),
            Build = new BuildOptions
            {
                PageLevel = pageLevel,
                BodyFontSizePt = bodyFontSize,
                Footer = Blank(Footer) ?? HelpDocument.DefaultFooter,
            },
            Compile = new CompileOptions
            {
                HhcPath = CompileChm ? Blank(HhcPath) : null,
            },
        };
    }

    /// <summary>
    /// Captures the inputs of a conversion. Compilation is considered requested when an
    /// hhc.exe was supplied, which is how <see cref="ConversionOptions"/> expresses it.
    /// </summary>
    public static ProjectFile FromConversionOptions(ConversionOptions options) => new()
    {
        DocxPath = options.DocxPath,
        OutputDirectory = options.OutputDirectory,
        BaseName = options.BaseName,
        ContextIdHeaderPath = options.ContextIdHeaderPath,
        ChmFileName = options.ChmFileName,
        ChmCopyPath = options.ChmCopyPath,
        PageLevel = options.Build.PageLevel,
        BodyFontSizePt = options.Build.BodyFontSizePt,
        Footer = options.Build.Footer,
        TemplateDirectory = options.TemplateDirectory,
        HhcPath = options.Compile.HhcPath,
        CompileChm = !string.IsNullOrWhiteSpace(options.Compile.HhcPath),
    };

    /// <summary>
    /// Writes the project to <paramref name="path"/>, adding the .w2c extension when the
    /// name has none. Paths that point next to the project file are made relative first.
    /// </summary>
    public void Save(string path)
    {
        var target = EnsureExtension(path);
        // Relative paths are only meaningful against the folder the file lands in, and that
        // may well be the folder the project is being copied into.
        var directory = Path.GetDirectoryName(Path.GetFullPath(target));

        MakePortable(directory);

        var json = JsonSerializer.Serialize(this, WriteOptions);

        try
        {
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(target, json);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            throw new ProjectFileException(
                $"Non è stato possibile salvare il progetto in \"{target}\": {ex.Message}", ex);
        }

        // On success the in-memory copy matches what is on disk, so it also becomes the
        // resolved location: a later Save writes back to the same file.
        SourcePath = Path.GetFullPath(target);
        MakeAbsolute(directory);
    }

    /// <summary>Reads a project file written by <see cref="Save"/>.</summary>
    public static ProjectFile Load(string path)
    {
        string text;
        try
        {
            text = File.ReadAllText(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            throw new ProjectFileException(
                $"Non è stato possibile leggere il progetto \"{path}\": {ex.Message}", ex);
        }

        ProjectFile? project;
        try
        {
            project = JsonSerializer.Deserialize<ProjectFile>(text, ReadOptions);
        }
        catch (JsonException ex)
        {
            throw new ProjectFileException(
                $"Il file \"{path}\" non è un progetto Word2Chm valido.", ex);
        }

        if (project is null)
        {
            throw new ProjectFileException($"Il file \"{path}\" non è un progetto Word2Chm valido.");
        }

        var fullPath = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(fullPath);

        // Relative paths were written against the folder the project was saved in; resolve
        // them against wherever the file is now, which is what makes the project portable.
        project.MakeAbsolute(directory);
        project.SourcePath = fullPath;
        project.Version = CurrentVersion;

        return project;
    }

    /// <summary>
    /// The JSON this project would be saved as, used to tell whether anything changed.
    /// Paths are compared as written, so the result does not depend on the current folder.
    /// </summary>
    public string Serialize() => JsonSerializer.Serialize(this, WriteOptions);

    /// <summary>Adds the .w2c extension when the given name does not already end with one.</summary>
    public static string EnsureExtension(string path)
    {
        if (Path.HasExtension(path) ||
            path.EndsWith(Path.DirectorySeparatorChar) ||
            path.EndsWith(Path.AltDirectorySeparatorChar))
        {
            return path;
        }

        return path + "." + Extension;
    }

    /// <summary>Default file name offered for a document, e.g. <c>guida.w2c</c>.</summary>
    public static string SuggestedFileName(string docxPath)
    {
        var name = Path.GetFileNameWithoutExtension(docxPath);
        return string.IsNullOrWhiteSpace(name) ? "progetto." + Extension : name + "." + Extension;
    }

    private static string? Blank(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    /// <summary>
    /// Paths that describe the document being worked on, so they are the ones worth making
    /// portable. hhc.exe belongs to the machine rather than to the project folder and stays
    /// absolute.
    /// </summary>
    private void MakePortable(string? directory)
    {
        DocxPath = Portable(DocxPath, directory);
        ContextIdHeaderPath = Portable(ContextIdHeaderPath, directory);
        OutputDirectory = Portable(OutputDirectory, directory);
        ChmCopyPath = Portable(ChmCopyPath, directory);
    }

    private void MakeAbsolute(string? directory)
    {
        DocxPath = Absolute(DocxPath, directory);
        ContextIdHeaderPath = Absolute(ContextIdHeaderPath, directory);
        OutputDirectory = Absolute(OutputDirectory, directory);
        ChmCopyPath = Absolute(ChmCopyPath, directory);
    }

    /// <summary>
    /// Rewrites an absolute path as one relative to <paramref name="directory"/> when it sits
    /// under a sibling folder, so the project folder can be moved as a whole. Paths outside
    /// that tree and relative paths are left untouched.
    /// </summary>
    private static string? Portable(string? path, string? directory)
    {
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrEmpty(directory) || !Path.IsPathRooted(path))
        {
            return path;
        }

        var relative = Path.GetRelativePath(directory, path);
        // A path that climbs a level or more is not part of the portable folder, so it is
        // kept absolute rather than written as a fragile "../..".
        return relative.StartsWith("..", StringComparison.Ordinal) ? path : relative;
    }

    private static string? Absolute(string? path, string? directory)
    {
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrEmpty(directory) || Path.IsPathRooted(path))
        {
            return path;
        }

        return Path.GetFullPath(Path.Combine(directory, path));
    }
}
