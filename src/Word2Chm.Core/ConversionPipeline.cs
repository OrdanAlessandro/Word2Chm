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

    /// <summary>Base name for the generated .chm/.hhp/.h files.</summary>
    public string? BaseName { get; init; }

    public BuildOptions Build { get; init; } = new();

    public CompileOptions Compile { get; init; } = new();

    /// <summary>Relative path stamped into the generated header comment.</summary>
    public string? ChmRelativePath { get; init; }
}

public sealed class ConversionResult
{
    public required HelpDocument Document { get; init; }
    public required IReadOnlyList<string> GeneratedFiles { get; init; }
    public required CompileResult Compilation { get; init; }
    public required string OutputDirectory { get; init; }
    public string? HeaderPath { get; init; }
    public string? ChmPath { get; init; }
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
            throw new FileNotFoundException("File DOCX non trovato.", options.DocxPath);
        }

        var outputDirectory = Path.GetFullPath(options.OutputDirectory);
        Directory.CreateDirectory(outputDirectory);

        var baseName = string.IsNullOrWhiteSpace(options.BaseName)
            ? Path.GetFileNameWithoutExtension(options.DocxPath)
            : options.BaseName!;
        var names = new ChmProjectNames { BaseName = baseName };

        var parsed = _parser.Parse(options.DocxPath);
        var document = _builder.Build(parsed, options.Build);

        var generated = new List<string>();
        var assetsDirectory = Path.Combine(outputDirectory, "assets");
        Directory.CreateDirectory(assetsDirectory);

        // Stylesheet.
        var cssPath = Path.Combine(outputDirectory, CssGenerator.FileName);
        File.WriteAllText(cssPath, CssGenerator.Generate());
        generated.Add(cssPath);

        // Assets.
        foreach (var image in parsed.Images)
        {
            var imagePath = Path.Combine(assetsDirectory, image.FileName);
            File.WriteAllBytes(imagePath, image.Content);
            generated.Add(imagePath);
        }

        // Pages plus a landing page that hosts the table of contents.
        foreach (var page in document.Pages)
        {
            var pagePath = Path.Combine(outputDirectory, page.FileName);
            File.WriteAllText(pagePath, _html.GeneratePage(page, document, CssGenerator.FileName));
            generated.Add(pagePath);
        }

        var indexPath = Path.Combine(outputDirectory, "index.html");
        File.WriteAllText(indexPath, _html.GenerateIndexPage(document));
        generated.Add(indexPath);

        // Help Workshop project files.
        var files = document.Pages.Select(p => p.FileName)
            .Concat(new[] { "index.html", CssGenerator.FileName })
            .Concat(parsed.Images.Select(i => "assets/" + i.FileName))
            .ToList();

        var hhpPath = Path.Combine(outputDirectory, names.HhpFile);
        File.WriteAllText(hhpPath, ChmProjectGenerator.GenerateHhp(document, names, files));
        generated.Add(hhpPath);

        var hhcPath = Path.Combine(outputDirectory, names.HhcFile);
        File.WriteAllText(hhcPath, ChmProjectGenerator.GenerateHhc(document));
        generated.Add(hhcPath);

        if (document.IndexEntries.Count > 0)
        {
            var hhkPath = Path.Combine(outputDirectory, names.HhkFile);
            File.WriteAllText(hhkPath, ChmProjectGenerator.GenerateHhk(document));
            generated.Add(hhkPath);
        }

        var headerPath = Path.Combine(outputDirectory, names.HeaderFile);
        File.WriteAllText(headerPath, ChmProjectGenerator.GenerateHeader(document, names, options.ChmRelativePath));
        generated.Add(headerPath);

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
            HeaderPath = headerPath,
            ChmPath = compilation.Success ? chmPath : null,
        };
    }
}
