using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Word2Chm.Core.Tests;

/// <summary>
/// Builds a representative .docx file in code so the conversion pipeline can be
/// exercised without shipping a binary fixture.
/// </summary>
internal static class DocxFixture
{
    public static byte[] CreateSample()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();

            // Numbering with one bulleted and one ordered list.
            var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = BuildNumbering();

            var body = new Body();
            main.Document = new Document(body);

            stylesPart.Styles = BuildStyles();
            stylesPart.Styles.Save();

            body.Append(Heading("Panoramica {#IDH_PANORAMICA}", 1));
            body.Append(TextParagraph(Run("Questa guida spiega come "), Run("installare"), Run(" il prodotto. ")));
            body.Append(HyperlinkParagraph(main));

            body.Append(Heading("Requisiti {#IDH_REQUISITI}", 1));
            body.Append(ListParagraph(1, 0, "Windows 10 o superiore"));
            body.Append(ListParagraph(1, 1, "2 GB di RAM"));
            body.Append(ListParagraph(2, 0, "Primo passo"));
            body.Append(ListParagraph(2, 0, "Secondo passo"));
            body.Append(IndexParagraph("requisiti", null, 1, "Requisiti di sistema"));
            body.Append(IndexParagraph("requisiti", "hardware", 1, "Requisiti hardware"));

            body.Append(Heading("Installazione {#IDH_INSTALLAZIONE}", 1));
            body.Append(Heading("Procedura {#IDH_PROCEDURA}", 2));
            body.Append(Heading("Configurazione città predefinita", 3));
            body.Append(TextParagraph(Run("Aprire il file "), Run("setup", bold: true), Run(" ed eseguire.")));
            body.Append(ImageParagraph(main, 320, 200));
            body.Append(CodeParagraph("dotnet build -c Release"));
            body.Append(BuildTable());
            body.Append(PageBreak());

            body.Append(Heading("Riferimenti {#IDH_RIFERIMENTI=5000}", 1));
            body.Append(TextParagraph(Run("Vedi anche la sezione ")));
            body.Append(BookmarkParagraph("cap-install"));

            main.Document.Save();
        }

        return stream.ToArray();
    }

    private static Paragraph Heading(string text, int level) => new(
        new ParagraphProperties(new ParagraphStyleId { Val = "Heading" + level }),
        new Run(new Text(text)));

    /// <summary>
    /// A document authored by an Italian Word: the heading style has a localised id
    /// ("Titolo1") and no explicit outline level, which is the shape that used to
    /// yield zero headings, an empty C++ header and a broken table of contents.
    /// The XE field is emitted the way Word really writes it - the instruction split
    /// across three runs inside nested begin/separate/end field characters - because
    /// reading only the first run was a second cause of the missing index.
    /// </summary>
    public static byte[] CreateLocalizedSample()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();

            var body = new Body();
            main.Document = new Document(body);

            var styles = new Styles();
            for (var level = 1; level <= 3; level++)
            {
                styles.Append(new Style(
                    new StyleName { Val = "Titolo " + level },
                    new BasedOn { Val = "Normal" })
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Titolo" + level,
                });
            }

            stylesPart.Styles = styles;
            stylesPart.Styles.Save();

            body.Append(LocalizedHeading("Capitolo primo {#IDH_PRIMO}", "Titolo1"));
            body.Append(TextParagraph(Run("Testo del capitolo.")));
            body.Append(SplitXeHeading("Sezione {#IDH_SEZIONE}", "Titolo2", "sezione"));
            body.Append(TextParagraph(Run("Testo della sezione.")));

            main.Document.Save();
        }

        return stream.ToArray();
    }

    private static Paragraph LocalizedHeading(string text, string styleId) => new(
        new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
        new Run(new Text(text)));

    /// <summary>
    /// A heading carrying an XE index field whose instruction is split across runs, the
    /// layout Word produces when it wraps the index entry in a bookmark.
    /// </summary>
    private static Paragraph SplitXeHeading(string text, string styleId, string keyword)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
            new Run(new Text(text)),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" XE \"") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldCode(keyword)),
            new Run(new FieldCode("\" ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        return paragraph;
    }

    private static Paragraph TextParagraph(params Run[] runs)
    {
        var paragraph = new Paragraph();
        foreach (var run in runs)
        {
            paragraph.Append(run);
        }

        return paragraph;
    }

    private static Run Run(string text, bool bold = false, bool italic = false)
    {
        var run = new Run();
        if (bold || italic)
        {
            var props = new RunProperties();
            if (bold)
            {
                props.Append(new Bold());
            }

            if (italic)
            {
                props.Append(new Italic());
            }

            run.Append(props);
        }

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static Paragraph HyperlinkParagraph(MainDocumentPart main)
    {
        var rel = main.AddHyperlinkRelationship(new Uri("https://example.com/docs"), true);
        return new Paragraph(
            new Run(new Text("Consulta la ")),
            new Hyperlink(new Run(new Text("documentazione"))) { Id = rel.Id, History = true },
            new Run(new Text(" ufficiale.")));
    }

    private static Paragraph BookmarkParagraph(string bookmarkName)
    {
        var paragraph = new Paragraph(
            new BookmarkStart { Id = "1", Name = bookmarkName },
            new Run(new Text("capitolo installazione")),
            new BookmarkEnd { Id = "1" });
        return paragraph;
    }

    private static Paragraph ListParagraph(int numId, int level, string text)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = level },
                    new NumberingId { Val = numId })),
            new Run(new Text(text)));
        return paragraph;
    }

    private static Paragraph IndexParagraph(string keyword, string? subKeyword, int numId, string text)
    {
        var instruction = subKeyword is null
            ? $"XE \"{keyword}\""
            : $"XE \"{keyword}\" \\t \"{subKeyword}\"";
        var field = new SimpleField { Instruction = instruction };
        field.Append(new Run(new Text(text)));
        return new Paragraph(
            new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = numId })),
            field);
    }

    private static Paragraph CodeParagraph(string text) => new(
        new ParagraphProperties(new ParagraphStyleId { Val = "Codice" }),
        new Run(new Text(text)));

    private static Paragraph ImageParagraph(MainDocumentPart main, int widthPx, int heightPx)
    {
        var imagePart = main.AddImagePart(ImagePartType.Png);
        using (var imageStream = imagePart.GetStream())
        {
            imageStream.Write(TinyPng);
        }

        var relId = main.GetIdOfPart(imagePart);
        const long emuPerPixel = 914400L / 96L;

        var drawing = new Drawing(
            new Wp.Inline(
                new Wp.Extent { Cx = widthPx * emuPerPixel, Cy = heightPx * emuPerPixel },
                new Wp.DocProperties { Id = (uint)1, Name = "Immagine" },
                new A.Graphic(
                    new A.GraphicData(
                        new Pic.Picture(
                            new Pic.NonVisualPictureProperties(
                                new Pic.NonVisualDrawingProperties { Id = 1U, Name = "image.png" },
                                new Pic.NonVisualPictureDrawingProperties()),
                            new Pic.BlipFill(
                                new A.Blip { Embed = relId },
                                new A.Stretch(new A.FillRectangle())),
                            new Pic.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = widthPx * emuPerPixel, Cy = heightPx * emuPerPixel }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

        return new Paragraph(new Run(drawing));
    }

    private static Table BuildTable() =>
        new(
            new TableRow(
                new TableRowProperties(new TableHeader()),
                new TableCell(new Paragraph(new Run(new Text("Colonna")))),
                new TableCell(new Paragraph(new Run(new Text("Valore"))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("Versione")))),
                new TableCell(new Paragraph(new Run(new Text("1.0"))))));

    private static Paragraph PageBreak() => new(new Run(new Break { Type = BreakValues.Page }));

    private static Styles BuildStyles()
    {
        var styles = new Styles();
        for (var level = 1; level <= 3; level++)
        {
            styles.Append(new Style(
                new StyleName { Val = "heading " + level },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new OutlineLevel { Val = level - 1 }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading" + level,
            });
        }

        styles.Append(new Style(
            new StyleName { Val = "Codice" },
            new StyleRunProperties(new RunFonts { Ascii = "Consolas" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "Codice",
        });

        return styles;
    }

    private static Numbering BuildNumbering()
    {
        var numbering = new Numbering();

        var bullet = new AbstractNum(
            new Level(
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = "•" })
            { LevelIndex = 0 });
        bullet.AbstractNumberId = 0;

        var ordered = new AbstractNum(
            new Level(
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." })
            { LevelIndex = 0 });
        ordered.AbstractNumberId = 1;

        numbering.Append(bullet);
        numbering.Append(ordered);
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 2 });
        return numbering;
    }

    /// <summary>Smallest valid PNG, used as an embedded image placeholder.</summary>
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
}
