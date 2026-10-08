using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Word2Chm.Core.Common;
using Word2Chm.Core.Compilation;
using Word2Chm.Core.Docx;
using Word2Chm.Core.Generation;
using Word2Chm.Core.Model;

namespace Word2Chm.Core;

public sealed class ConversionOptions
{
    public required string DocxPath { get; init; }

    /// <summary>Output directory; created when missing.</summary>
    public required string OutputDirectory { get; init; }

    /// <summary>Base name for the generated .hhp/.hhc/.hhk project files.</summary>
    public string? BaseName { get; init; }

    /// <summary>
    /// C++ header that supplies the numeric context IDs, e.g. <c>helpId.h</c>. The document
    /// names a symbol and the number is read from here, so one header can back every
    /// language edition and the IDs stay identical across all of them.
    /// </summary>
    public string? ContextIdHeaderPath { get; init; }

    /// <summary>
    /// Name of the compiled help file (default <c>SkipperQtHelp_IT.chm</c>). Only the file
    /// name is used, so the .chm always lands next to the other generated files.
    /// </summary>
    public string? ChmFileName { get; init; }

    public BuildOptions Build { get; init; } = new();

    public CompileOptions Compile { get; init; } = new();

    /// <summary>
    /// Directory holding the WinCHM skin (fixedtop.htm and its css/js/gif). When set and
    /// the template file exists, every topic is wrapped in the skin.
    /// </summary>
    public string? TemplateDirectory { get; init; }
}

public sealed class ConversionResult
{
    public required HelpDocument Document { get; init; }
    public required IReadOnlyList<string> GeneratedFiles { get; init; }
    public required CompileResult Compilation { get; init; }
    public required string OutputDirectory { get; init; }
    public string? ChmPath { get; init; }

    /// <summary>
    /// Text file listing the symbols the document used but the header did not define. Null
    /// when every symbol resolved; the file is written even so the CHM is produced.
    /// </summary>
    public string? MissingIdsReportPath { get; init; }
}

/// <summary>
/// End-to-end conversion: parse DOCX, build the help project, write HTML/CSS/assets,
/// emit the Help Workshop files and the C++ header, then optionally compile the CHM.
/// </summary>
public sealed class ConversionPipeline
{
    private readonly DocxParser _parser = new();
    private readonly HelpProjectBuilder _builder = new();
    private readonly HtmlGenerator _html = new();
    private readonly ChmCompiler _compiler = new();

    public ConversionResult Run(ConversionOptions options)
    {
        if (!File.Exists(options.DocxPath))
        {
            throw new DocxUnreadableException(
                $"Il documento \"{options.DocxPath}\" non esiste o non è raggiungibile.");
        }

        var outputDirectory = Path.GetFullPath(options.OutputDirectory);
        Directory.CreateDirectory(outputDirectory);

        var baseName = string.IsNullOrWhiteSpace(options.BaseName)
            ? Path.GetFileNameWithoutExtension(options.DocxPath)
            : options.BaseName!;
        var names = new ChmProjectNames
        {
            BaseName = baseName,
            ChmFileName = options.ChmFileName,
        };

        var parsed = ParseDocument(options.DocxPath);

        // The header supplied by the user is the source of truth for the numeric IDs. A
        // missing or unreadable header is a mistake to fix before converting, not something
        // to work around silently, so it is reported as such.
        var build = options.Build;
        build.ContextIds = LoadContextIdHeader(options.ContextIdHeaderPath);
        var document = _builder.Build(parsed, build);

        var generated = new List<string>();
        var assetsDirectory = Path.Combine(outputDirectory, "assets");
        Directory.CreateDirectory(assetsDirectory);

        // Stylesheet.
        var cssPath = Path.Combine(outputDirectory, CssGenerator.FileName);
        File.WriteAllText(cssPath, CssGenerator.Generate(options.Build.BodyFontSizePt));
        generated.Add(cssPath);

        // WinCHM skin: the template plus its stylesheet, script and button images.
        var templateFiles = new List<string>();
        var template = LoadTemplate(options.TemplateDirectory, outputDirectory, generated, templateFiles);

        // Assets.
        foreach (var image in parsed.Images)
        {
            var imagePath = Path.Combine(assetsDirectory, image.FileName);
            File.WriteAllBytes(imagePath, image.Content);
            generated.Add(imagePath);
        }

        // Pages plus a landing page that hosts the table of contents. Previous/next links
        // follow the table-of-contents order so the reader can page through the help.
        var order = FlattenPages(document);
        for (var index = 0; index < order.Count; index++)
        {
            var page = order[index];
            var previous = index > 0 ? order[index - 1].FileName : null;
            var next = index + 1 < order.Count ? order[index + 1].FileName : null;

            var pagePath = Path.Combine(outputDirectory, page.FileName);
            File.WriteAllText(pagePath, _html.GeneratePage(page, document, CssGenerator.FileName, template, previous, next));
            generated.Add(pagePath);
        }

        // Redirect topics for context IDs declared on sub-headings. The compiler resolves an
        // alias only to a topic file, so each anchored ID gets a stub that forwards to the
        // anchor; these files must also be listed in [FILES].
        var anchorFiles = new List<string>();
        foreach (var page in document.Pages)
        {
            foreach (var anchor in page.Anchors)
            {
                var stubPath = Path.Combine(outputDirectory, anchor.FileName);
                File.WriteAllText(stubPath, _html.GenerateAnchorPage(anchor, page.Title, document));
                generated.Add(stubPath);
                anchorFiles.Add(anchor.FileName);
            }
        }

        var indexPath = Path.Combine(outputDirectory, "index.html");
        File.WriteAllText(indexPath, _html.GenerateIndexPage(document));
        generated.Add(indexPath);

        // Help Workshop project files. The index must also be declared in [FILES],
        // otherwise its keywords never make it into the compiled CHM.
        var hasIndex = document.IndexEntries.Count > 0;
        var files = document.Pages.Select(p => p.FileName)
            .Concat(new[] { "index.html", CssGenerator.FileName })
            .Concat(anchorFiles)
            .Concat(hasIndex ? new[] { names.HhkFile } : Array.Empty<string>())
            .Concat(parsed.Images.Select(i => "assets/" + i.FileName))
            .Concat(templateFiles)
            .ToList();

        var hhpPath = Path.Combine(outputDirectory, names.HhpFile);
        File.WriteAllText(hhpPath, ChmProjectGenerator.GenerateHhp(document, names, files));
        generated.Add(hhpPath);

        var hhcPath = Path.Combine(outputDirectory, names.HhcFile);
        File.WriteAllText(hhcPath, ChmProjectGenerator.GenerateHhc(document));
        generated.Add(hhcPath);

        if (hasIndex)
        {
            var hhkPath = Path.Combine(outputDirectory, names.HhkFile);
            File.WriteAllText(hhkPath, ChmProjectGenerator.GenerateHhk(document));
            generated.Add(hhkPath);
        }

        // Symbols the document used but the header does not define. The CHM is still built so
        // the rest of the work is not lost; the list is written to a text file the UI opens.
        string? missingReportPath = null;
        if (build.MissingSymbols.Count > 0)
        {
            missingReportPath = Path.Combine(outputDirectory, "id-mancanti.txt");
            File.WriteAllText(missingReportPath, BuildMissingIdsReport(build.MissingSymbols, options.ContextIdHeaderPath));
            generated.Add(missingReportPath);
        }

        var chmPath = Path.Combine(outputDirectory, names.ChmFile);
        var compilation = _compiler.Compile(hhpPath, chmPath, options.Compile);

        if (compilation.Success && File.Exists(chmPath))
        {
            generated.Add(chmPath);
        }

        return new ConversionResult
        {
            Document = document,
            GeneratedFiles = generated,
            Compilation = compilation,
            OutputDirectory = outputDirectory,
            ChmPath = compilation.Success ? chmPath : null,
            MissingIdsReportPath = missingReportPath,
        };
    }

    /// <summary>
    /// Reads the C++ header with the ID definitions. A blank path yields an empty set, which
    /// makes every symbol "missing" and is reported; a path pointing at a file that is not
    /// there is a mistake to fix, so it stops the conversion.
    /// </summary>
    private static ContextIdHeader LoadContextIdHeader(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return ContextIdHeader.Parse(string.Empty);
        }

        if (!File.Exists(path))
        {
            throw new DocxUnreadableException(
                $"Il file degli ID \"{path}\" non esiste o non è raggiungibile.");
        }

        try
        {
            return ContextIdHeader.Load(path);
        }
        catch (IOException ex)
        {
            throw new DocxUnreadableException(
                $"Il file degli ID \"{path}\" non è leggibile: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Writes the list of symbols missing from the header, ready to be pasted into it. A
    /// numeric placeholder is included so the user only has to replace the numbers, which
    /// is faster than typing every <c>#define</c> by hand.
    /// </summary>
    private static string BuildMissingIdsReport(IReadOnlyList<string> missing, string? headerPath)
    {
        var builder = new StringBuilder();
        builder.AppendLine("ID di contesto mancanti nel file .h");
        builder.AppendLine("====================================");
        builder.AppendLine();
        builder.AppendLine(headerPath is not null
            ? $"File .h utilizzato: {headerPath}"
            : "File .h utilizzato: (nessuno)");
        builder.AppendLine();
        builder.AppendLine(
            $"I seguenti {missing.Count} simboli sono usati nel documento Word ma non sono definiti nel file .h:");
        builder.AppendLine("aggiungi le righe indicate al file .h e ripeti la conversione.");
        builder.AppendLine();
        foreach (var symbol in missing)
        {
            builder.AppendLine($"  {symbol}");
        }

        builder.AppendLine();
        builder.AppendLine("Righe da aggiungere al file .h (sostituisci i valori numerici):");
        builder.AppendLine();
        foreach (var symbol in missing)
        {
            builder.AppendLine($"#define {symbol,-40} 0");
        }

        return builder.ToString();
    }

    /// <summary>
    /// Reads the .docx, turning the ways an unusable file fails into one exception the UI
    /// can present. Word keeps the document open with a lock, so a file that exists can
    /// still be unreadable; a renamed .docx or a wrong file throws from the Open XML SDK.
    /// </summary>
    private ParsedDocument ParseDocument(string path)
    {
        try
        {
            return _parser.Parse(path);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new DocxUnreadableException(
                $"Il documento \"{path}\" è in uso o non è accessibile. " +
                "Chiudilo in Word e riprova.", ex);
        }
        catch (IOException ex)
        {
            throw new DocxUnreadableException(
                $"Impossibile leggere il documento \"{path}\": {ex.Message}", ex);
        }
        catch (FileFormatException ex)
        {
            throw new DocxUnreadableException(
                $"Il file \"{path}\" non è un documento Word (.docx) valido.", ex);
        }
        catch (Exception ex) when (ex is not DocxUnreadableException)
        {
            throw new DocxUnreadableException(
                $"Impossibile aprire il documento \"{path}\": {ex.Message}", ex);
        }
    }

    private static List<HelpPage> FlattenPages(HelpDocument document)
    {
        var byFile = document.Pages.ToDictionary(p => p.FileName, StringComparer.Ordinal);
        var ordered = new List<HelpPage>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        void Visit(TocNode node)
        {
            // A table-of-contents entry may point at "page.html#anchor" while the page is
            // keyed by "page.html", so the key used for de-duplication is the file name.
            // Keying on the raw Local put every page in the order twice, which made the
            // previous/next buttons jump backwards (the first page linked to the last).
            if (!string.IsNullOrEmpty(node.Local))
            {
                var file = node.Local.Split('#')[0];
                if (seen.Add(file) && byFile.TryGetValue(file, out var page))
                {
                    ordered.Add(page);
                }
            }

            foreach (var child in node.Children)
            {
                Visit(child);
            }
        }

        foreach (var node in document.Toc)
        {
            Visit(node);
        }

        // Pages never reached through the table of contents keep their document order.
        foreach (var page in document.Pages)
        {
            if (seen.Add(page.FileName))
            {
                ordered.Add(page);
            }
        }

        return ordered;
    }

    /// <summary>
    /// Reads the skin and copies it into the output next to the topics. Returns null when no
    /// template directory was given, so the built-in layout stays available.
    /// </summary>
    private static string? LoadTemplate(
        string? templateDirectory,
        string outputDirectory,
        List<string> generated,
        List<string> templateFiles)
    {
        if (string.IsNullOrWhiteSpace(templateDirectory))
        {
            return null;
        }

        var templatePath = Path.Combine(templateDirectory, WinChmTemplate.TemplateFileName);
        if (!File.Exists(templatePath))
        {
            return null;
        }

        var source = WinChmTemplate.ReadTemplate(templatePath);
        foreach (var name in new[]
                 {
                     WinChmTemplate.StyleFileName,
                     WinChmTemplate.ScriptFileName,
                 }.Concat(WinChmTemplate.ImageFileNames))
        {
            var from = Path.Combine(templateDirectory, name);
            if (!File.Exists(from))
            {
                continue;
            }

            File.Copy(from, Path.Combine(outputDirectory, name), overwrite: true);
            generated.Add(Path.Combine(outputDirectory, name));
            templateFiles.Add(name);
        }

        return source;
    }
}
