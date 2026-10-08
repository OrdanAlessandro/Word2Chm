using Word2Chm.Core.Generation;
using Word2Chm.Core.Model;
using Word2Chm.Core.Project;

namespace Word2Chm.Core.Tests;

public sealed class ProjectFileTests : IDisposable
{
    private readonly string _workDirectory;

    public ProjectFileTests()
    {
        _workDirectory = Path.Combine(Path.GetTempPath(), "w2c-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_workDirectory);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_workDirectory, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private static ProjectFile Sample() => new()
    {
        DocxPath = "/tmp/non-importa/guida.docx",
        OutputDirectory = "/tmp/non-importa/out",
        BaseName = "guida",
        ContextIdHeaderPath = "/tmp/non-importa/ids.h",
        ChmFileName = "manuale.chm",
        ChmCopyPath = "/tmp/non-importa/dist",
        PageLevel = 4,
        BodyFontSizePt = 11,
        Footer = "Pie di pagina",
        TemplateDirectory = "/tmp/non-importa/skin",
        HhcPath = "/tmp/non-importa/hhc.exe",
        CompileChm = true,
        OpenOutputOnFinish = true,
    };

    [Fact]
    public void RoundTripsEverySetting()
    {
        var path = Path.Combine(_workDirectory, "progetto.w2c");
        Sample().Save(path);

        var loaded = ProjectFile.Load(path);

        Assert.Equal("guida", loaded.BaseName);
        Assert.Equal("manuale.chm", loaded.ChmFileName);
        Assert.Equal(4, loaded.PageLevel);
        Assert.Equal(11, loaded.BodyFontSizePt);
        Assert.Equal("Pie di pagina", loaded.Footer);
        Assert.True(loaded.CompileChm);
        Assert.True(loaded.OpenOutputOnFinish);
    }

    [Fact]
    public void AddsTheExtensionWhenTheNameHasNone()
    {
        var path = Path.Combine(_workDirectory, "senzaestensione");
        Sample().Save(path);

        Assert.True(File.Exists(path + ".w2c"));
        Assert.Equal(path + ".w2c", ProjectFile.Load(path + ".w2c").SourcePath);
    }

    [Fact]
    public void SavesDocumentAndHeaderRelativeToTheProjectFile()
    {
        // The usual layout has the document, the header and the project in one folder, so
        // those paths are written relative and survive the folder being moved.
        var directory = Path.Combine(_workDirectory, "portatile");
        Directory.CreateDirectory(directory);
        var docx = Path.Combine(directory, "guida.docx");
        var header = Path.Combine(directory, "ids.h");
        File.WriteAllText(docx, "x");
        File.WriteAllText(header, "x");

        var project = Sample();
        project.DocxPath = docx;
        project.ContextIdHeaderPath = header;
        project.OutputDirectory = Path.Combine(directory, "out");
        project.ChmCopyPath = Path.Combine(directory, "dist", "manuale.chm");

        var path = Path.Combine(directory, "guida.w2c");
        project.Save(path);

        var json = File.ReadAllText(path);
        Assert.Contains("\"guida.docx\"", json);
        Assert.Contains("\"ids.h\"", json);
        Assert.DoesNotContain(directory, json);

        // Loading resolves them against the folder the file is in, so the values come back
        // absolute and usable.
        var loaded = ProjectFile.Load(path);
        Assert.Equal(docx, loaded.DocxPath);
        Assert.Equal(header, loaded.ContextIdHeaderPath);
        Assert.Equal(Path.Combine(directory, "out"), loaded.OutputDirectory);
        Assert.Equal(Path.Combine(directory, "dist", "manuale.chm"), loaded.ChmCopyPath);
    }

    [Fact]
    public void ResolvesRelativePathsAfterTheFolderIsMoved()
    {
        var original = Path.Combine(_workDirectory, "originale");
        Directory.CreateDirectory(original);
        File.WriteAllText(Path.Combine(original, "guida.docx"), "x");

        var project = Sample();
        project.DocxPath = Path.Combine(original, "guida.docx");
        project.ContextIdHeaderPath = null;
        project.Save(Path.Combine(original, "guida.w2c"));

        var moved = Path.Combine(_workDirectory, "spostato");
        Directory.Move(original, moved);

        // The document moved with the project, so the relative path still points at it even
        // though the absolute location changed.
        var loaded = ProjectFile.Load(Path.Combine(moved, "guida.w2c"));
        Assert.Equal(Path.Combine(moved, "guida.docx"), loaded.DocxPath);
    }

    [Fact]
    public void KeepsPathsOutsideTheProjectFolderAbsolute()
    {
        // A document that lives somewhere else is not part of the portable folder, so it is
        // written as it is rather than as a fragile relative path climbing out of the folder.
        var outside = Path.Combine(_workDirectory, "altrove", "guida.docx");
        var projectDirectory = Path.Combine(_workDirectory, "progetto");
        Directory.CreateDirectory(Path.GetDirectoryName(outside)!);
        Directory.CreateDirectory(projectDirectory);

        var project = Sample();
        project.DocxPath = outside;
        project.ContextIdHeaderPath = null;
        var path = Path.Combine(projectDirectory, "guida.w2c");
        project.Save(path);

        Assert.Contains(outside.Replace("\\", "\\\\"), File.ReadAllText(path));
        Assert.Equal(outside, ProjectFile.Load(path).DocxPath);
    }

    [Fact]
    public void SavingTwiceTargetsTheSameFile()
    {
        // After the first save the project knows where it lives, so a plain "save" must not
        // need the path again nor create a second file.
        var path = Path.Combine(_workDirectory, "progetto.w2c");
        var project = Sample();
        project.Save(path);
        project.BaseName = "rinominato";
        project.Save(project.SourcePath!);

        var files = Directory.GetFiles(_workDirectory, "*.w2c");
        Assert.Single(files);
        Assert.Equal("rinominato", ProjectFile.Load(path).BaseName);
    }

    [Fact]
    public void ReportsAFileThatIsNotAProject()
    {
        var path = Path.Combine(_workDirectory, "rotto.w2c");
        File.WriteAllText(path, "{ questo non e' json");

        Assert.Throws<ProjectFileException>(() => ProjectFile.Load(path));
    }

    [Fact]
    public void ClampsValuesOutsideTheirRange()
    {
        var path = Path.Combine(_workDirectory, "fuori.w2c");
        File.WriteAllText(path, """
            { "pageLevel": 99, "bodyFontSizePt": 200 }
            """);

        var options = ProjectFile.Load(path).ToConversionOptions();

        Assert.Equal(BuildOptions.MaxPageLevel, options.Build.PageLevel);
        Assert.Equal(BuildOptions.DefaultBodyFontSizePt, options.Build.BodyFontSizePt);
    }

    [Fact]
    public void CompilesOnlyWhenCompilationIsRequested()
    {
        var project = Sample();
        project.CompileChm = false;

        var options = project.ToConversionOptions();

        // The path is still stored, so switching compilation back on restores the same
        // hhc.exe, but nothing is passed to the pipeline while it is off.
        Assert.Null(options.Compile.HhcPath);
        Assert.Equal("/tmp/non-importa/hhc.exe", project.HhcPath);
    }

    [Fact]
    public void SuggestedNameFollowsTheDocument()
    {
        Assert.Equal("guida.w2c", ProjectFile.SuggestedFileName("/tmp/guida.docx"));
        Assert.Equal("progetto.w2c", ProjectFile.SuggestedFileName(string.Empty));
    }
}
