using Word2Chm.Core;
using Word2Chm.Core.Compilation;
using Word2Chm.Core.Generation;

namespace Word2Chm.Core.Tests;

public sealed class ConversionPipelineTests : IDisposable
{
    private readonly string _workDirectory;

    public ConversionPipelineTests()
    {
        _workDirectory = Path.Combine(Path.GetTempPath(), "word2chm-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_workDirectory);
    }

    public void Dispose()
    {
        if (Directory.Exists(_workDirectory))
        {
            Directory.Delete(_workDirectory, recursive: true);
        }
    }

    private string WriteSampleDocx()
    {
        var path = Path.Combine(_workDirectory, "guida.docx");
        File.WriteAllBytes(path, DocxFixture.CreateSample());
        return path;
    }

    private ConversionResult RunPipeline(string? hhcPath = null, string? baseName = null)
    {
        var pipeline = new ConversionPipeline();
        return pipeline.Run(new ConversionOptions
        {
            DocxPath = WriteSampleDocx(),
            OutputDirectory = Path.Combine(_workDirectory, "out"),
            BaseName = baseName,
            Build = new BuildOptions { DefaultContextId = 1000 },
            Compile = new CompileOptions { HhcPath = hhcPath },
        });
    }

    [Fact]
    public void CreatesOnePagePerLevelOneHeading()
    {
        var result = RunPipeline();

        Assert.Equal(4, result.Document.Pages.Count);
        Assert.Equal("Panoramica", result.Document.Pages[0].Title);
        Assert.Equal("Requisiti", result.Document.Pages[1].Title);
        Assert.Equal("Installazione", result.Document.Pages[2].Title);
        Assert.Equal("Riferimenti", result.Document.Pages[3].Title);
    }

    [Fact]
    public void StripsContextIdMarkerFromVisibleTitle()
    {
        var result = RunPipeline();

        Assert.All(result.Document.Pages, page => Assert.DoesNotContain("{#", page.Title));
        Assert.Equal("IDH_INSTALLAZIONE", result.Document.Pages[2].Symbol);
    }

    [Fact]
    public void AssignsSequentialContextIdsButHonoursExplicitOnes()
    {
        var result = RunPipeline();

        Assert.Equal(1000, result.Document.Pages[0].ContextId);
        Assert.Equal(1001, result.Document.Pages[1].ContextId);
        Assert.Equal(1002, result.Document.Pages[2].ContextId);

        // The fourth page declares {#IDH_RIFERIMENTI=5000}.
        Assert.Equal(5000, result.Document.Pages[3].ContextId);
    }

    [Fact]
    public void GeneratesHeaderWithDefineForEachSymbol()
    {
        var result = RunPipeline();

        var header = File.ReadAllText(result.HeaderPath!);
        Assert.Contains("#define IDH_PANORAMICA", header);
        Assert.Contains("#define IDH_INSTALLAZIONE", header);
        Assert.Contains("#define IDH_RIFERIMENTI", header);
        Assert.Contains("5000", header);
    }

    [Fact]
    public void HhpContainsMapAndAliasSections()
    {
        var result = RunPipeline();
        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));

        Assert.Contains("[MAP]", hhp);
        Assert.Contains("[ALIAS]", hhp);
        Assert.Contains("IDH_INSTALLAZIONE=002-installazione.html", hhp);
        Assert.Contains("#define IDH_INSTALLAZIONE", hhp);
    }

    [Fact]
    public void ExtractsImagesAndWritesThemToAssets()
    {
        var result = RunPipeline();
        var assets = Path.Combine(result.OutputDirectory, "assets");

        Assert.True(Directory.Exists(assets));
        Assert.NotEmpty(Directory.GetFiles(assets));
    }

    [Fact]
    public void ProducesIndexEntriesFromXeFields()
    {
        var result = RunPipeline();

        Assert.Contains(result.Document.IndexEntries, e => e.Keyword == "requisiti" && e.SubKeyword == null);
        Assert.Contains(result.Document.IndexEntries, e => e.Keyword == "requisiti" && e.SubKeyword == "hardware");
    }

    [Fact]
    public void GeneratesHhkWhenIndexEntriesExist()
    {
        var result = RunPipeline();
        var hhkPath = Path.Combine(result.OutputDirectory, "guida.hhk");

        Assert.True(File.Exists(hhkPath));
        var hhk = File.ReadAllText(hhkPath);
        Assert.Contains("requisiti", hhk);
        Assert.Contains("hardware", hhk);
    }

    [Fact]
    public void HhkKeepsPlainAndSubKeywordEntries()
    {
        var result = RunPipeline();
        var hhk = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhk"));

        // The keyword has both a plain target and a "hardware" sub-keyword; both must survive.
        Assert.Equal(2, CountOccurrences(hhk, "name=\"Name\" value=\"requisiti\""));
        Assert.Contains("name=\"Name\" value=\"hardware\"", hhk);
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        var count = 0;
        var index = 0;
        while ((index = haystack.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += needle.Length;
        }

        return count;
    }

    [Fact]
    public void ResolvesExternalHyperlinkInHtml()
    {
        var result = RunPipeline();
        var firstPage = File.ReadAllText(Path.Combine(result.OutputDirectory, result.Document.Pages[0].FileName));

        Assert.Contains("https://example.com/docs", firstPage);
    }

    [Fact]
    public void BuildsTableOfContentsHierarchy()
    {
        var result = RunPipeline();

        Assert.NotEmpty(result.Document.Toc);
        var installazione = result.Document.Toc.Single(n => n.Title == "Installazione");
        Assert.Contains(installazione.Children, n => n.Title == "Procedura");
    }

    [Fact]
    public void RendersTableAndListMarkup()
    {
        var result = RunPipeline();
        var installPage = File.ReadAllText(Path.Combine(result.OutputDirectory, result.Document.Pages[2].FileName));
        var requisitiPage = File.ReadAllText(Path.Combine(result.OutputDirectory, result.Document.Pages[1].FileName));

        Assert.Contains("<table>", installPage);
        Assert.Contains("<th>", installPage);
        Assert.Contains("<hr class=\"pagebreak\">", installPage);
        Assert.Contains("<ol", requisitiPage); // numId 2 is the ordered list
        Assert.Contains("<ul", requisitiPage); // numId 1 is the bullet list
    }

    [Fact]
    public void TrimsWhitespaceLeftByTheContextIdMarker()
    {
        var result = RunPipeline();
        var installPage = File.ReadAllText(Path.Combine(result.OutputDirectory, result.Document.Pages[2].FileName));

        Assert.Contains("<h1 id=\"installazione\">Installazione</h1>", installPage);
    }

    [Fact]
    public void ReportsFailureWhenHhcIsMissing()
    {
        var result = RunPipeline();

        Assert.False(result.Compilation.Success);
        Assert.Null(result.ChmPath);
        Assert.Contains("hhc.exe", result.Compilation.Output);
    }

    [Fact]
    public void FailsWhenHhcPathDoesNotExist()
    {
        var result = RunPipeline(hhcPath: Path.Combine(_workDirectory, "nope.exe"));

        Assert.False(result.Compilation.Success);
    }

    [Fact]
    public void UsesFallbackTitleFromFileNameWhenNoCoreProperties()
    {
        var result = RunPipeline();
        Assert.Equal("guida", result.Document.Title);
    }

    [Fact]
    public void ThrowsForMissingInput()
    {
        var pipeline = new ConversionPipeline();
        Assert.Throws<FileNotFoundException>(() => pipeline.Run(new ConversionOptions
        {
            DocxPath = Path.Combine(_workDirectory, "missing.docx"),
            OutputDirectory = Path.Combine(_workDirectory, "out"),
        }));
    }
}
