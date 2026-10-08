using System.Text;
using Word2Chm.Core;
using Word2Chm.Core.Common;
using Word2Chm.Core.Compilation;
using Word2Chm.Core.Generation;
using Word2Chm.Core.Model;

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

    /// <summary>
    /// Writes a header defining <paramref name="symbols"/> with sequential numbers, mirroring
    /// the file the user supplies. It is written once per test run and reused.
    /// </summary>
    private string WriteContextIdHeader(params string[] symbols)
    {
        var path = Path.Combine(_workDirectory, "ids-" + Guid.NewGuid().ToString("N") + ".h");
        var builder = new StringBuilder();
        builder.AppendLine("#pragma once");
        builder.AppendLine();
        var id = 1000;
        foreach (var symbol in symbols)
        {
            builder.AppendLine($"#define {symbol} {id++}");
        }

        File.WriteAllText(path, builder.ToString());
        return path;
    }

    private ConversionResult RunPipeline(
        string? hhcPath = null,
        string? baseName = null,
        string? templateDirectory = null,
        int? pageLevel = null,
        string? contextIdHeaderPath = null,
        string? chmFileName = null)
    {
        // Every symbol the sample document declares, so a plain RunPipeline resolves them all.
        var header = contextIdHeaderPath ?? WriteContextIdHeader(
            "IDH_PANORAMICA", "IDH_REQUISITI", "IDH_INSTALLAZIONE", "IDH_PROCEDURA",
            "IDH_RIFERIMENTI", "IDH_CAPITOLO", "IDH_PRIMO", "IDH_SECONDO",
            "IDH_OPERAZIONI", "IDH_SEZIONE");

        var pipeline = new ConversionPipeline();
        var build = new BuildOptions();
        if (pageLevel is not null)
        {
            // Left unset otherwise, so the tests follow the production default instead of
            // pinning a level that could drift away from it.
            build.PageLevel = pageLevel.Value;
        }

        return pipeline.Run(new ConversionOptions
        {
            DocxPath = WriteSampleDocx(),
            OutputDirectory = Path.Combine(_workDirectory, "out"),
            BaseName = baseName,
            ContextIdHeaderPath = header,
            ChmFileName = chmFileName,
            Build = build,
            TemplateDirectory = templateDirectory,
            Compile = new CompileOptions { HhcPath = hhcPath },
        });
    }

    /// <summary>Runs the pipeline over a document whose headings carry no body text.</summary>
    private ConversionResult RunHeadingOnlyPipeline(int pageLevel)
    {
        var path = Path.Combine(_workDirectory, "vuoto.docx");
        File.WriteAllBytes(path, DocxFixture.CreateHeadingOnlySample());

        return new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "out-vuoto"),
            BaseName = "guida",
            Build = new BuildOptions { PageLevel = pageLevel },
            Compile = new CompileOptions { HhcPath = null },
        });
    }

    /// <summary>Writes the fixedtop skin into the test directory and returns its folder.</summary>
    private string WriteTemplate()
    {
        var directory = Path.Combine(_workDirectory, "skin");
        Directory.CreateDirectory(directory);

        File.WriteAllText(Path.Combine(directory, "fixedtop.htm"),
            "\uFEFF<html><head><link href=\"winchm_template_style.css\" rel=\"stylesheet\">" +
            "<script src=\"winchm_template_script.js\"></script></head><body>" +
            "<div id=\"top\"><img src=\"btn_prev_n.gif\"><img src=\"btn_next_n.gif\"></div>" +
            "<div id=\"nav\">($navigation$)</div><div id=\"title\">($title$)</div>" +
            "<div id=\"content\">($content$)</div><div id=\"footer\">($footer$)</div></body></html>");
        File.WriteAllText(Path.Combine(directory, "winchm_template_style.css"), "body{}");
        File.WriteAllText(Path.Combine(directory, "winchm_template_script.js"), "//");
        File.WriteAllBytes(Path.Combine(directory, "btn_prev_n.gif"), new byte[] { 0x47, 0x49, 0x46 });
        File.WriteAllBytes(Path.Combine(directory, "btn_next_n.gif"), new byte[] { 0x47, 0x49, 0x46 });

        return directory;
    }

    [Fact]
    public void CreatesOnePagePerHeading()
    {
        // The default page level splits every heading, not just the level-one ones, so the
        // sample yields six topics in document order.
        var result = RunPipeline();

        Assert.Equal(6, result.Document.Pages.Count);
        Assert.Equal(
            new[] { "Panoramica", "Requisiti", "Installazione", "Procedura", "Configurazione città predefinita", "Riferimenti" },
            result.Document.Pages.Select(p => p.Title));
    }

    [Fact]
    public void TakesContextIdsFromTheSuppliedHeader()
    {
        // The numbers come from the header, not from the order in the document, so the same
        // header keeps the IDs identical across language editions.
        var header = Path.Combine(_workDirectory, "mio.h");
        File.WriteAllText(header, """
            #pragma once
            #define IDH_PANORAMICA   4001
            #define IDH_REQUISITI    4002
            #define IDH_INSTALLAZIONE 0x0FDA
            #define IDH_PROCEDURA    4004
            #define IDH_RIFERIMENTI  4005
            #define IDH_CAPITOLO     4006
            #define IDH_PRIMO        4007
            #define IDH_SECONDO      4008
            #define IDH_OPERAZIONI   4009
            #define IDH_SEZIONE      4010
            """);

        var result = RunPipeline(contextIdHeaderPath: header);

        Assert.Equal(4001, result.Document.Pages.Single(p => p.Title == "Panoramica").ContextId);
        Assert.Equal(4002, result.Document.Pages.Single(p => p.Title == "Requisiti").ContextId);
        // The header value is in hexadecimal, proving the number is parsed rather than guessed.
        Assert.Equal(0x0FDA, result.Document.Pages.Single(p => p.Title == "Installazione").ContextId);
    }

    [Fact]
    public void IgnoresTheLegacyInlineNumber()
    {
        // The sample declares {#IDH_RIFERIMENTI=5000} on the last page. That number is now
        // meaningless: the header is the only source, because a number in the document could
        // not stay consistent across languages.
        var result = RunPipeline();

        Assert.Equal(1004, result.Document.Pages.Single(p => p.Title == "Riferimenti").ContextId);
    }

    [Fact]
    public void ReportsSymbolsMissingFromTheHeaderAndStillBuilds()
    {
        // Only two symbols are defined; the others must be listed, and the conversion must
        // still complete so the CHM and the rest of the work are not lost.
        var header = WriteContextIdHeader("IDH_PANORAMICA", "IDH_INSTALLAZIONE");

        var result = RunPipeline(contextIdHeaderPath: header);

        Assert.NotNull(result.MissingIdsReportPath);
        var report = File.ReadAllText(result.MissingIdsReportPath!);
        Assert.Contains("IDH_RIFERIMENTI", report);
        Assert.Contains("IDH_PROCEDURA", report);
        Assert.DoesNotContain("IDH_PANORAMICA", report);

        // The symbol that is defined keeps its ID, the missing ones get none.
        Assert.Equal(1000, result.Document.Pages.Single(p => p.Title == "Panoramica").ContextId);
        Assert.Null(result.Document.Pages.Single(p => p.Title == "Riferimenti").ContextId);

        // The CHM was still produced.
        Assert.True(File.Exists(Path.Combine(result.OutputDirectory, "guida.hhp")));
    }

    [Fact]
    public void WritesNoMapEntryForAMissingSymbol()
    {
        var header = WriteContextIdHeader("IDH_PANORAMICA");

        var result = RunPipeline(contextIdHeaderPath: header);

        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));
        Assert.Contains("#define IDH_PANORAMICA 1000", hhp);
        Assert.DoesNotContain("#define IDH_RIFERIMENTI", hhp);
    }

    [Fact]
    public void FailsWhenTheHeaderFileDoesNotExist()
    {
        var error = Assert.Throws<DocxUnreadableException>(() =>
            RunPipeline(contextIdHeaderPath: Path.Combine(_workDirectory, "nope.h")));

        Assert.Contains("non esiste", error.Message);
    }

    [Fact]
    public void StripsContextIdMarkerFromVisibleTitle()
    {
        var result = RunPipeline();

        Assert.All(result.Document.Pages, page => Assert.DoesNotContain("{#", page.Title));
        Assert.Equal("IDH_INSTALLAZIONE", result.Document.Pages.Single(p => p.Title == "Installazione").Symbol);
    }

    [Fact]
    public void NamesTheChmSkipperQtHelpByDefault()
    {
        // The host application loads the help file by name, so the default cannot follow
        // the base name: a conversion of "guida.docx" still produces SkipperQtHelp_IT.chm.
        var result = RunPipeline(baseName: "guida");

        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));
        Assert.Contains("Compiled file=SkipperQtHelp_IT.chm", hhp);
    }

    [Fact]
    public void UsesTheConfiguredChmName()
    {
        var result = RunPipeline(chmFileName: "manuale.chm");

        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));
        Assert.Contains("Compiled file=manuale.chm", hhp);
    }

    [Fact]
    public void AppendsTheChmExtensionWhenItIsMissing()
    {
        // Typing "manuale" in the UI should still produce a .chm.
        var result = RunPipeline(chmFileName: "manuale");

        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));
        Assert.Contains("Compiled file=manuale.chm", hhp);
    }

    [Fact]
    public void FallsBackToTheDefaultChmNameWhenBlank()
    {
        var result = RunPipeline(chmFileName: "   ");

        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));
        Assert.Contains("Compiled file=SkipperQtHelp_IT.chm", hhp);
    }

    [Fact]
    public void KeepsTheChmInsideTheOutputDirectory()
    {
        // The .hhp compiled-file value must stay a plain file name, otherwise hhc.exe would
        // write the .chm somewhere else (or fail) depending on the working directory.
        var result = RunPipeline(chmFileName: "../fuori.chm");

        var hhp = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));
        Assert.Contains("Compiled file=.._fuori.chm", hhp);
        Assert.DoesNotContain("Compiled file=../", hhp);
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

        // The table and the page break sit under the level-three heading, which now has a
        // topic of its own.
        var configPage = ReadPage(result, "Configurazione città predefinita");
        var requisitiPage = ReadPage(result, "Requisiti");

        Assert.Contains("<table>", configPage);
        Assert.Contains("<th>", configPage);
        Assert.Contains("<hr class=\"pagebreak\">", configPage);
        Assert.Contains("<ol", requisitiPage); // numId 2 is the ordered list
        Assert.Contains("<ul", requisitiPage); // numId 1 is the bullet list
    }

    [Fact]
    public void TrimsWhitespaceLeftByTheContextIdMarker()
    {
        var result = RunPipeline();
        var installPage = ReadPage(result, "Installazione");

        Assert.Contains("<h1 id=\"installazione\">Installazione</h1>", installPage);
    }

    /// <summary>Reads the generated HTML of the topic whose heading has this title.</summary>
    private static string ReadPage(ConversionResult result, string title)
    {
        var page = result.Document.Pages.Single(p => p.Title == title);
        return File.ReadAllText(Path.Combine(result.OutputDirectory, page.FileName));
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
            ContextIdHeaderPath = WriteContextIdHeader("IDH_PRIMO", "IDH_SEZIONE"),
            Build = new BuildOptions { PageLevel = 1 },
            Compile = new CompileOptions(),
        });

        Assert.Single(result.Document.Pages);
        Assert.Equal("Capitolo primo", result.Document.Pages[0].Title);
        Assert.Equal("IDH_PRIMO", result.Document.Pages[0].Symbol);

        var hhp0 = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp"));
        Assert.Contains("#define IDH_PRIMO", hhp0);

        // IDH_SEZIONE sits on a level-2 heading. The compiler resolves an alias only to a
        // topic file ("file.htm#anchor" triggers HHC3015 and the CHM then fails to open),
        // so the symbol targets a small redirect topic that forwards to the section.
        Assert.Contains("#define IDH_SEZIONE", hhp0);

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
            Build = new BuildOptions(),
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

    /// <summary>
    /// The default page level must match the table of contents: every entry the menu shows
    /// is a topic of its own. Sharing one page across several entries is what makes the
    /// viewer display a whole chapter as a single long document.
    /// </summary>
    [Fact]
    public void SplitsOnePagePerTableOfContentsEntryByDefault()
    {
        var result = RunPipeline();

        var toc = new List<TocNode>();
        Flatten(result.Document.Toc, toc);

        Assert.Equal(toc.Count, result.Document.Pages.Count);
        foreach (var group in toc.Where(n => n.Local is not null)
                                 .GroupBy(n => n.Local!.Split('#')[0], StringComparer.Ordinal))
        {
            Assert.Single(group);
        }
    }

    /// <summary>
    /// A section label that expands into sub-headings keeps its own topic. Merging it into
    /// the preceding page used to point its menu entry at the tail of an unrelated section.
    /// </summary>
    [Fact]
    public void KeepsItsOwnPageForASectionLabelWithChildren()
    {
        var path = Path.Combine(_workDirectory, "contenitore.docx");
        File.WriteAllBytes(path, DocxFixture.CreateSectionLabelSample());

        var result = new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "out-contenitore"),
            BaseName = "guida",
            Build = new BuildOptions(),
            Compile = new CompileOptions { HhcPath = null },
        });

        var label = result.Document.Pages.Single(p => p.Title == "Utensili");

        // The label page is not the page of the chapter that precedes it.
        var first = result.Document.Pages.Single(p => p.Title == "Prima sezione");
        Assert.NotEqual(first.FileName, label.FileName);

        // ...and the child heading still has a page of its own.
        Assert.Contains(result.Document.Pages, p => p.Title == "Fresa");
    }

    private static void Flatten(IEnumerable<TocNode> nodes, List<TocNode> into)
    {
        foreach (var node in nodes)
        {
            into.Add(node);
            Flatten(node.Children, into);
        }
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
        var error = Assert.Throws<DocxUnreadableException>(() => pipeline.Run(new ConversionOptions
        {
            DocxPath = Path.Combine(_workDirectory, "missing.docx"),
            OutputDirectory = Path.Combine(_workDirectory, "out"),
        }));

        Assert.Contains("non esiste o non è raggiungibile", error.Message);
    }

    /// <summary>
    /// The navigation pane is a Win32 tree view, so "Window Styles" carries tree-view bits:
    /// TVS_HASBUTTONS (1) draws the +/- boxes, TVS_HASLINES (2) the connecting lines and
    /// TVS_LINESATROOT (4) the lines to the root items. The value 0x23520 cleared all three
    /// and set TVS_CHECKBOXES (0x100), which is the reported symptom.
    /// </summary>
    [Fact]
    public void HhcEnablesTreeLinesInsteadOfCheckBoxes()
    {
        var result = RunPipeline();
        var hhc = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhc"));

        var match = System.Text.RegularExpressions.Regex.Match(
            hhc, "<param name=\"Window Styles\" value=\"(0x[0-9a-fA-F]+)\">");
        Assert.True(match.Success, "Window Styles non impostato in .hhc");

        const int hasButtons = 0x1, hasLines = 0x2, linesAtRoot = 0x4, checkBoxes = 0x100;
        var styles = Convert.ToInt32(match.Groups[1].Value, 16);
        Assert.True((styles & hasButtons) != 0, $"TVS_HASBUTTONS mancante: {styles:X}");
        Assert.True((styles & hasLines) != 0, $"TVS_HASLINES mancante: {styles:X}");
        Assert.True((styles & linesAtRoot) != 0, $"TVS_LINESATROOT mancante: {styles:X}");
        Assert.True((styles & checkBoxes) == 0, $"TVS_CHECKBOXES impostato: {styles:X}");
    }

    [Fact]
    public void WrapsPagesInTemplateAndCopiesItsAssets()
    {
        var templateDirectory = WriteTemplate();
        var result = RunPipeline(templateDirectory: templateDirectory);

        var page = File.ReadAllText(Path.Combine(result.OutputDirectory, "001-requisiti.html"));

        // Placeholders are gone and the page content sits inside the skin.
        Assert.DoesNotContain("($title$)", page);
        Assert.DoesNotContain("($content$)", page);
        Assert.Contains("winchm_template_top", page.Replace("id=\"top\"", "id=\"winchm_template_top\""));
        Assert.Contains("Requisiti", page);

        // The skin assets end up next to the topics and are declared in [FILES],
        // otherwise the buttons are missing from the compiled CHM.
        foreach (var name in new[]
                 {
                     "winchm_template_style.css", "winchm_template_script.js",
                     "btn_prev_n.gif", "btn_next_n.gif",
                 })
        {
            Assert.True(File.Exists(Path.Combine(result.OutputDirectory, name)), $"{name} mancante");
            Assert.Contains(name, File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhp")));
        }

        // The project stylesheet is still linked, or tables and code lose their formatting.
        Assert.Contains("help.css", page);
    }

    [Fact]
    public void TemplateButtonsLinkNeighbouringPages()
    {
        var result = RunPipeline(templateDirectory: WriteTemplate());
        var pages = result.Document.Pages;

        var first = File.ReadAllText(Path.Combine(result.OutputDirectory, pages[0].FileName));
        var second = File.ReadAllText(Path.Combine(result.OutputDirectory, pages[1].FileName));

        // The first page has no previous topic, the second links both ways.
        Assert.DoesNotContain($"<a href=\"\"><img src=\"btn_prev_n.gif\">", first);
        Assert.Contains($"btn_next_n.gif", first);
        Assert.Contains($"href=\"{pages[0].FileName}\"", second);
        Assert.Contains($"href=\"{pages[2].FileName}\"", second);
    }

    [Fact]
    public void BreadcrumbLinksAncestorHeadings()
    {
        // With pages split at level 2, "Procedura" is a topic under "Installazione", so its
        // page carries a trail that links back to the parent heading.
        var result = RunPipeline(templateDirectory: WriteTemplate(), pageLevel: 2);

        var page = result.Document.Pages.Single(p => p.Title == "Procedura");
        Assert.Single(page.Ancestors);
        Assert.Equal("Installazione", page.Ancestors[0].Title);

        var html = File.ReadAllText(Path.Combine(result.OutputDirectory, page.FileName));
        Assert.Contains("href=\"" + page.Ancestors[0].Local + "\"", html);
        Assert.Contains("Installazione", html);
    }

    /// <summary>
    /// A heading that carries no text of its own is a structural label, not a topic. Demoting
    /// it keeps the menu free of one-line pages while its heading stays as an anchor on the
    /// parent page, so nothing that links to it breaks.
    /// </summary>
    [Fact]
    public void DemotesHeadingOnlyPagesIntoTheirParent()
    {
        var result = RunHeadingOnlyPipeline(pageLevel: 2);

        Assert.DoesNotContain(result.Document.Pages, p => p.Title == "Sezione vuota");
        Assert.Contains(result.Document.Pages, p => p.Title == "Sezione piena");

        // The demoted heading still appears in the output, as an anchor of the parent page.
        var parent = result.Document.Pages.Single(p => p.Title == "Capitolo");
        var html = File.ReadAllText(Path.Combine(result.OutputDirectory, parent.FileName));
        Assert.Contains("id=\"sezione-vuota\"", html);

        // ...and the table of contents keeps pointing at a file that exists.
        var hhc = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhc"));
        Assert.Contains(parent.FileName + "#sezione-vuota", hhc);
    }

    [Fact]
    public void KeepsPageForHeadingCarryingAContextId()
    {
        // A context ID is a promise to the C++ header and the [ALIAS] section, so the page
        // must survive even when the heading has no body text.
        var result = RunHeadingOnlyPipeline(pageLevel: 2);
        Assert.Contains(result.Document.Pages, p => p.Symbol == "IDH_CAPITOLO");
        Assert.NotNull(result.Document.Pages.Single(p => p.Symbol == "IDH_CAPITOLO").Blocks);
    }

    [Fact]
    public void DoesNotLinkToRemovedPagesAfterDemotion()
    {
        var result = RunHeadingOnlyPipeline(pageLevel: 2);
        var existing = result.Document.Pages.Select(p => p.FileName).ToHashSet(StringComparer.Ordinal);

        var hhc = File.ReadAllText(Path.Combine(result.OutputDirectory, "guida.hhc"));
        foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(
                     hhc, "<param name=\"Local\" value=\"([^\"]+)\""))
        {
            var file = match.Groups[1].Value.Split('#')[0];
            Assert.True(existing.Contains(file), $"il sommario punta a un file rimosso: {file}");
        }

        foreach (var page in result.Document.Pages)
        {
            var html = File.ReadAllText(Path.Combine(result.OutputDirectory, page.FileName));
            foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(
                         html, "href=\"([^\"]+)\""))
            {
                var href = match.Groups[1].Value;
                if (href.StartsWith("http", StringComparison.OrdinalIgnoreCase) || href.StartsWith('#'))
                {
                    continue;
                }

                var file = href.Split('#')[0];
                Assert.True(string.IsNullOrEmpty(file) || existing.Contains(file) ||
                            File.Exists(Path.Combine(result.OutputDirectory, file)),
                    $"collegamento rotto in {page.FileName}: {href}");
            }
        }
    }

    [Fact]
    public void ContinuesNumberingWhenAListIsInterruptedByText()
    {
        // Regression: Word keeps one counter per list definition for the whole document, so
        // the run after the interrupting paragraph must start at 3. The parser emitted a
        // separate <ol> per run, and every one of them restarted at 1.
        var path = Path.Combine(_workDirectory, "elenco.docx");
        File.WriteAllBytes(path, DocxFixture.CreateResumedListSample());

        var result = new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "elenco-out"),
            BaseName = "guida",
            Build = new BuildOptions { PageLevel = 1 },
            Compile = new CompileOptions(),
        });

        var page = result.Document.Pages.Single(p => p.Title == "Primo capitolo");
        var html = File.ReadAllText(Path.Combine(result.OutputDirectory, page.FileName));

        var lists = System.Text.RegularExpressions.Regex.Matches(html, "<ol[^>]*>");
        Assert.Equal(2, lists.Count);
        Assert.Contains("start=\"3\"", lists[1].Value);

        // Word's own marker format has to survive as well, or an ordered list renders as a
        // browser default that ignores the numbering Word assigned.
        Assert.Contains("list-style-type:decimal", html);
    }

    [Fact]
    public void KeepsOneListWhenBlankParagraphsSeparateTheItems()
    {
        // Regression: Word writes an empty paragraph between the items of a list. Those
        // separators used to close the run, so the chapter "Operazioni preliminari all'invio"
        // came out as one <ol> per item, each rendering as 1 in a viewer that ignores start.
        var path = Path.Combine(_workDirectory, "elenco-vuoti.docx");
        File.WriteAllBytes(path, DocxFixture.CreateListWithBlankSeparatorsSample());

        var result = new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "elenco-vuoti-out"),
            BaseName = "guida",
            Build = new BuildOptions { PageLevel = 1 },
            Compile = new CompileOptions(),
        });

        var page = result.Document.Pages.Single(p => p.Title == "Operazioni preliminari");
        var html = File.ReadAllText(Path.Combine(result.OutputDirectory, page.FileName));

        var lists = System.Text.RegularExpressions.Regex.Matches(html, "<ol[^>]*>");
        Assert.Single(lists);

        // One list means the items number themselves natively, with no reliance on start.
        Assert.Equal(3, System.Text.RegularExpressions.Regex.Matches(html, "<li").Count);
        Assert.DoesNotContain("start=", html);
    }

    [Fact]
    public void WritesTheConfiguredBodyFontSizeToTheStylesheet()
    {
        // Regression: the body size was hardcoded at 10.5pt, so the setting had no effect
        // and the WinCHM skin kept overriding the content div with its own 8.5pt.
        var path = Path.Combine(_workDirectory, "font.docx");
        File.WriteAllBytes(path, DocxFixture.CreateSample());

        var result = new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "font-out"),
            BaseName = "guida",
            Build = new BuildOptions { PageLevel = 1, BodyFontSizePt = 14 },
            Compile = new CompileOptions(),
        });

        var css = File.ReadAllText(Path.Combine(result.OutputDirectory, "help.css"));

        Assert.Contains("--body-font-size: 14pt;", css);
        Assert.Contains("font-size: var(--body-font-size);", css);

        // The skin sizes its content div with an id selector, which outranks body.
        Assert.Contains("#winchm_template_content { font-size: var(--body-font-size); }", css);

        // Headings stay relative, so enlarging the body must not change them.
        Assert.Contains("h1 { font-size: 1.9em;", css);
    }

    [Fact]
    public void DefaultsTheBodyFontSizeWhenNotConfigured()
    {
        var path = Path.Combine(_workDirectory, "font-default.docx");
        File.WriteAllBytes(path, DocxFixture.CreateSample());

        var result = new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "font-default-out"),
            BaseName = "guida",
            Build = new BuildOptions { PageLevel = 1 },
            Compile = new CompileOptions(),
        });

        var css = File.ReadAllText(Path.Combine(result.OutputDirectory, "help.css"));
        Assert.Contains("--body-font-size: 10.5pt;", css);
    }

    [Fact]
    public void ResolvesCrossReferenceToBookmarkWrittenBesideAParagraph()
    {
        // Regression: Word writes a cross-reference target as a direct child of the body,
        // between two paragraphs. Reading only the bookmarks inside a paragraph dropped it,
        // and the link fell back to href="#".
        var path = Path.Combine(_workDirectory, "riferimento.docx");
        File.WriteAllBytes(path, DocxFixture.CreateResumedListSample());

        var result = new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "rif-out"),
            BaseName = "guida",
            Build = new BuildOptions { PageLevel = 1 },
            Compile = new CompileOptions(),
        });

        var target = result.Document.Pages.Single(p => p.Title == "Secondo capitolo");
        Assert.True(result.Document.BookmarkTargets.TryGetValue("_Capitolo_due", out var resolved));
        Assert.Equal(target.FileName, resolved.PageFileName);

        var source = result.Document.Pages.Single(p => p.Title == "Primo capitolo");
        var html = File.ReadAllText(Path.Combine(result.OutputDirectory, source.FileName));

        Assert.Contains($"{target.FileName}#{resolved.Anchor}", html);
        Assert.DoesNotContain("href=\"#\"", html);
    }

    [Fact]
    public void KeepsItalicDeclaredByAParagraphStyle()
    {
        // Regression: Word italicises captions through the paragraph style, leaving the runs
        // bare. Reading only the run's own w:i dropped every caption in the document.
        var path = Path.Combine(_workDirectory, "corsivo-paragrafo.docx");
        File.WriteAllBytes(path, DocxFixture.CreateStyleItalicSample());

        var result = RunItalicPipeline(path);
        var html = ReadPage(result, "Didascalie");

        Assert.Contains("<em>Figura 1: pannello dei parametri</em>", html);
    }

    [Fact]
    public void KeepsItalicDeclaredByACharacterStyle()
    {
        // Regression: the "Emphasis" character style states the italic, not the run.
        var path = Path.Combine(_workDirectory, "corsivo-carattere.docx");
        File.WriteAllBytes(path, DocxFixture.CreateStyleItalicSample());

        var result = RunItalicPipeline(path);
        var html = ReadPage(result, "Enfasi");

        Assert.Contains("<em>enfasi</em>", html);
    }

    [Fact]
    public void CharacterStyleCancelsItalicInheritedFromTheParagraph()
    {
        // Regression: in an already italic caption the "Emphasis" style means the opposite
        // of what it means in body text. Word stores "Quando vengono modificati..." with a
        // Didascalia paragraph and an Enfasi run, and renders it upright: treating the style
        // as a plain "italic = true" wrongly italicised 154 paragraphs of the manual.
        var path = Path.Combine(_workDirectory, "corsivo-annullato.docx");
        File.WriteAllBytes(path, DocxFixture.CreateStyleItalicSample());

        var result = RunItalicPipeline(path);
        var html = ReadPage(result, "DidascalieEnfasi");

        Assert.Contains("Quando vengono modificati dei parametri", html);
        Assert.DoesNotContain("<em>Quando vengono modificati dei parametri</em>", html);
    }

    [Fact]
    public void ResolvesItalicThroughTheBasedOnChain()
    {
        // A style that states no italic of its own inherits it from the style it is based on.
        var path = Path.Combine(_workDirectory, "corsivo-ereditato.docx");
        File.WriteAllBytes(path, DocxFixture.CreateStyleItalicSample());

        var result = RunItalicPipeline(path);
        var html = ReadPage(result, "Ereditarieta");

        Assert.Contains("<em>Testo corsivo per ereditarieta.</em>", html);
    }

    [Fact]
    public void LetsADirectRunSettingSwitchItalicOff()
    {
        // w:i val="0" on the run means "not italic" even when the character style would
        // turn the italic on, so the explicit run value has to win.
        var path = Path.Combine(_workDirectory, "corsivo-disattivato.docx");
        File.WriteAllBytes(path, DocxFixture.CreateStyleItalicSample());

        var result = RunItalicPipeline(path);
        var html = ReadPage(result, "Disattivazione");

        Assert.Contains("<em>corsivo </em>non corsivo", html);
    }

    private ConversionResult RunItalicPipeline(string path) =>
        new ConversionPipeline().Run(new ConversionOptions
        {
            DocxPath = path,
            OutputDirectory = Path.Combine(_workDirectory, "corsivo-out-" + Guid.NewGuid().ToString("N")),
            BaseName = "guida",
            Build = new BuildOptions { PageLevel = 1 },
            Compile = new CompileOptions(),
        });

    [Fact]
    public void ReportsAFileThatIsNotAWordDocument()
    {
        // Regression: renaming any file to .docx used to leak FileFormatException from the
        // Open XML SDK straight into the popup. It is reported as an unusable document.
        var notDocx = Path.Combine(_workDirectory, "finto.docx");
        File.WriteAllText(notDocx, "questo non è un documento Word");

        var error = Assert.Throws<DocxUnreadableException>(() => RunItalicPipeline(notDocx));

        Assert.Contains("non è un documento Word", error.Message);
    }

    [Fact]
    public void ReportsADirectoryPassedAsTheDocument()
    {
        var directory = Path.Combine(_workDirectory, "cartella.docx");
        Directory.CreateDirectory(directory);

        var error = Assert.Throws<DocxUnreadableException>(() => RunItalicPipeline(directory));

        Assert.Contains("non esiste o non è raggiungibile", error.Message);
    }
}
