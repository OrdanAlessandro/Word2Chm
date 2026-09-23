using System.Text;
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

    /// <summary>
    /// The viewer reads the [WINDOWS] value positionally, so WindowStyles must stay at
    /// field 9. A stray comma moves it onto the window-rect field and the CHM then fails
    /// to open with "There is not enough memory available for this task".
    /// </summary>
    [Fact]
    public void HhpKeepsWindowsFieldsAligned()
    {
        var result = RunPipeline();
        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));

        var line = hhp.Split('\n')
            .SkipWhile(l => !l.Trim().Equals("[WINDOWS]", StringComparison.Ordinal))
            .Skip(1)
            .First(l => l.Contains('='));
        var fields = SplitWindowsFields(line[(line.IndexOf('=') + 1)..]);

        Assert.Equal("0x23520", fields[9]);
        Assert.Empty(fields[12]);
    }

    /// <summary>Splits a [WINDOWS] value on commas that sit outside double quotes.</summary>
    private static string[] SplitWindowsFields(string value)
    {
        var fields = new List<string>();
        var current = new StringBuilder();
        var quoted = false;
        foreach (var ch in value)
        {
            if (ch == '"')
            {
                quoted = !quoted;
                current.Append(ch);
            }
            else if (ch == ',' && !quoted)
            {
                fields.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(ch);
            }
        }

        fields.Add(current.ToString());
        return fields.ToArray();
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
    public void HhkEmitsEachKeywordOnceAndLinksTheParent()
    {
        var result = RunPipeline();
        var hhk = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhk"));

        // A keyword must appear exactly once as a node; duplicating it as a sibling
        // leaf produces a CHM that hh.exe refuses to open.
        Assert.Equal(1, CountOccurrences(hhk, "name=\"Name\" value=\"requisiti\""));
        Assert.Contains("name=\"Name\" value=\"hardware\"", hhk);

        // The "requisiti" node is itself a link, so the plain target is not lost.
        Assert.Contains("name=\"Local\" value=\"001-requisiti.html#requisiti\"", hhk);
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
    public void EncodesAccentedCharactersAsEntitiesInProjectFiles()
    {
        var result = RunPipeline();
        var directory = result.OutputDirectory;

        // Raw UTF-8 in .hhc/.hhk makes hhc.exe produce a CHM that will not open, so
        // these files must stay pure ASCII.
        Assert.True(IsPureAscii(File.ReadAllText(Path.Combine(directory, "guida.hhc"))));
        Assert.True(IsPureAscii(File.ReadAllText(Path.Combine(directory, "guida.hhk"))));
        Assert.True(IsPureAscii(File.ReadAllText(Path.Combine(directory, "guida.hhp"))));
    }

    private static bool IsPureAscii(string text) => text.All(ch => ch <= 127);

    [Fact]
    public void RecognizesHeadingsFromLocalizedStyleNames()
    {
        // Regression: a document whose heading style is "Titolo1" (Italian Word) and
        // carries no outline level used to produce no pages, an empty .h and a broken
        // table of contents, and the resulting CHM would not open.
        var path = Path.Combine(_workDirectory, "localizzato.docx");
        File.WriteAllBytes(path, DocxFixture.CreateLocalizedSample());

        var result = new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "loc-out"),
            BaseName = "guida",
            Build = new BuildOptions { DefaultContextId = 1000 },
            Compile = new CompileOptions(),
        });

        Assert.Single(result.Document.Pages);
        Assert.Equal("Capitolo primo", result.Document.Pages[0].Title);
        Assert.Equal("IDH_PRIMO", result.Document.Pages[0].Symbol);

        var header = File.ReadAllText(result.HeaderPath!);
        Assert.Contains("#define IDH_PRIMO", header);

        // IDH_SEZIONE sits on a level-2 heading. The compiler resolves an alias only to a
        // topic file ("file.htm#anchor" triggers HHC3015 and the CHM then fails to open),
        // so the symbol targets a small redirect topic that forwards to the section.
        Assert.Contains("#define IDH_SEZIONE", header);

        var page = result.Document.Pages[0];
        var anchor = Assert.Single(page.Anchors);
        Assert.Equal("IDH_SEZIONE", anchor.Symbol);
        Assert.Equal("sezione", anchor.Anchor);
        Assert.Equal(page.FileName + "#sezione", anchor.Target);

        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));

        // No alias may carry an anchor, or hhc.exe reports the file as missing.
        Assert.DoesNotContain("#", AliasSection(hhp));
        Assert.Contains("IDH_SEZIONE=" + anchor.FileName, hhp);

        var stub = Path.Combine(result.OutputDirectory, anchor.FileName);
        Assert.True(File.Exists(stub));
        Assert.Contains(page.FileName + "#sezione", File.ReadAllText(stub));
        Assert.Contains(anchor.FileName, hhp);
    }

    /// <summary>Returns the [ALIAS] section of a .hhp file.</summary>
    private static string AliasSection(string hhp)
    {
        var start = hhp.IndexOf("[ALIAS]", StringComparison.Ordinal);
        Assert.True(start >= 0);
        var end = hhp.IndexOf("\n[", start + 1, StringComparison.Ordinal);
        return end < 0 ? hhp[start..] : hhp[start..end];
    }

    [Fact]
    public void ReadsXeInstructionSplitAcrossRuns()
    {
        // Word writes an XE instruction as several runs (" XE \"" + keyword + "\" ").
        // Reading only the first FieldCode yielded no index at all.
        var path = Path.Combine(_workDirectory, "localizzato.docx");
        File.WriteAllBytes(path, DocxFixture.CreateLocalizedSample());

        var result = new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "loc-out"),
            BaseName = "guida",
            Build = new BuildOptions { DefaultContextId = 1000 },
            Compile = new CompileOptions(),
        });

        var entry = Assert.Single(result.Document.IndexEntries);
        Assert.Equal("sezione", entry.Keyword);
        Assert.Equal(result.Document.Pages[0].FileName, entry.FileName);

        // The index must also be declared in [FILES], or its keywords never reach the CHM.
        var hhkPath = Path.Combine(result.OutputDirectory, "guida.hhk");
        Assert.True(File.Exists(hhkPath));
        Assert.Contains("name=\"Name\" value=\"sezione\"", File.ReadAllText(hhkPath));
        Assert.Contains("guida.hhk", File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp")));
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
