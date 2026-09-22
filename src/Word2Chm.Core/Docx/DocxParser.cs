using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Word2Chm.Core.Common;
using Word2Chm.Core.Model;
using ModelTableCell = Word2Chm.Core.Model.TableCell;
using ModelTableRow = Word2Chm.Core.Model.TableRow;
using ModelListItem = Word2Chm.Core.Model.ListItem;
using OpenXmlTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using OpenXmlTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using OpenXmlListItem = DocumentFormat.OpenXml.Wordprocessing.ListItem;
using OpenXmlTableHeader = DocumentFormat.OpenXml.Wordprocessing.TableHeader;

namespace Word2Chm.Core.Docx;

/// <summary>
/// Reads a .docx package into the flat <see cref="DocumentModel"/> representation.
/// The parser resolves relationships (hyperlinks, images) and numbering metadata so
/// that the generation stage never has to touch the Open XML object model.
/// </summary>
public sealed class DocxParser
{
    private static readonly Regex XeFieldRegex = new(
        @"XE\s+""(?<main>[^""]*)""(?:\s*\\t\s*""(?<sub>[^""]*)""|\s*\\t\s*""(?<sub2>[^""]*)""|\s*\\t\s*(?<sub3>""[^""]*""))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex StyleHeadingRegex = new(
        @"^(?:Heading|Titolo|Titre|Überschrift|Título|Titel|Rubrik|Otsikko|Nadpis|Nagłówek)\s*(?<level>[1-9])$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public ParsedDocument Parse(string path)
    {
        using var stream = File.OpenRead(path);
        return Parse(stream, Path.GetFileNameWithoutExtension(path));
    }

    public ParsedDocument Parse(Stream stream, string fallbackTitle = "Guida")
    {
        var result = new ParsedDocument();
        var imageCache = new Dictionary<string, ImageResource>(StringComparer.Ordinal);
        var orderedImages = new List<ImageResource>();

        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart
            ?? throw new InvalidOperationException("Il documento non contiene una parte principale.");

        var numberingMap = BuildNumberingMap(main);
        var styleMap = BuildStyleMap(main);

        result.Title = ReadCoreTitle(document) ?? fallbackTitle;

        var body = main.Document.Body;
        if (body is null)
        {
            return result;
        }

        foreach (var element in body.ChildElements)
        {
            switch (element)
            {
                case Paragraph paragraph:
                    AddParagraph(paragraph, result, main, numberingMap, styleMap, imageCache, orderedImages);
                    break;
                case Table table:
                    result.Blocks.Add(ParseTable(table, main, numberingMap, styleMap, imageCache, orderedImages));
                    break;
            }
        }

        result.Images.AddRange(orderedImages);
        return result;
    }

    // ---------------------------------------------------------------- paragraphs

    private void AddParagraph(
        Paragraph paragraph,
        ParsedDocument result,
        MainDocumentPart main,
        IReadOnlyDictionary<int, List<NumberingLevel>> numberingMap,
        IReadOnlyDictionary<string, StyleInfo> styleMap,
        Dictionary<string, ImageResource> imageCache,
        List<ImageResource> orderedImages)
    {
        var props = paragraph.ParagraphProperties;
        var styleId = props?.ParagraphStyleId?.Val?.Value;
        var styleName = !string.IsNullOrEmpty(styleId) && styleMap.TryGetValue(styleId!, out var known)
            ? known.Name
            : null;
        var outline = ResolveOutlineLevel(props, styleId, styleMap);
        var headingLevel = ResolveHeadingLevel(styleId, styleName, outline);
        var bookmarks = paragraph.Descendants<BookmarkStart>().ToList();

        if (headingLevel > 0)
        {
            var heading = BuildHeading(paragraph, headingLevel, bookmarks, main, imageCache, orderedImages);
            if (heading is not null)
            {
                result.Blocks.Add(heading);
            }

            return;
        }

        if (IsListParagraph(props))
        {
            var list = BuildList(paragraph, numberingMap, main, imageCache, orderedImages);
            if (list is not null)
            {
                result.Blocks.Add(list);
                return;
            }
        }

        var indexKeywords = ExtractIndexKeywords(paragraph);
        var inlines = BuildInlines(paragraph, main, imageCache, orderedImages, out var lastImage, out var lastImageSize);

        if (IsCodeStyle(styleId, styleMap))
        {
            result.Blocks.Add(new CodeBlock
            {
                Text = string.Concat(inlines.OfType<TextInline>().Select(t => t.Text)),
                Language = null,
            });
            return;
        }

        if (inlines.Count == 0 && lastImage is not null)
        {
            result.Blocks.Add(new ImageBlock
            {
                FileName = lastImage,
                Width = lastImageSize.Width,
                Height = lastImageSize.Height,
            });
            return;
        }

        if (ContainsPageBreak(paragraph) && IsBlankApartFromBreaks(inlines))
        {
            result.Blocks.Add(new PageBreakBlock());
            return;
        }

        if (inlines.Count == 0 && indexKeywords.Count == 0 && !ContainsPageBreak(paragraph))
        {
            if (bookmarks.Count == 0)
            {
                return;
            }
        }

        var block = new ParagraphBlock { StyleId = styleId };
        block.IndexKeywords.AddRange(indexKeywords);
        block.Inlines.AddRange(inlines);
        AddBookmarkNames(block, bookmarks);

        if (inlines.Count == 1 && inlines[0] is TextInline { ImageFileName: not null } onlyImage)
        {
            block.ImageFileName = onlyImage.ImageFileName;
        }

        result.Blocks.Add(block);
    }

    private HeadingBlock? BuildHeading(
        Paragraph paragraph,
        int level,
        List<BookmarkStart> bookmarks,
        MainDocumentPart main,
        Dictionary<string, ImageResource> imageCache,
        List<ImageResource> orderedImages)
    {
        var inlines = BuildInlines(paragraph, main, imageCache, orderedImages, out _, out _);
        var text = InlineText.Flatten(inlines);
        var marker = ContextIdMarker.Parse(text);

        var heading = new HeadingBlock
        {
            Level = level,
            Title = marker.CleanText,
            Symbol = string.IsNullOrEmpty(marker.Symbol) ? null : marker.Symbol,
            ExplicitId = marker.ExplicitId,
        };

        heading.Inlines.AddRange(RewriteInlineText(inlines, marker));
        AddBookmarkNames(heading, bookmarks);

        return heading.Title.Length == 0 && heading.Inlines.Count == 0 ? null : heading;
    }

    private static void AddBookmarkNames(DocumentBlock block, IEnumerable<BookmarkStart> bookmarks)
    {
        foreach (var bookmark in bookmarks)
        {
            var name = bookmark.Name?.Value;
            if (!string.IsNullOrEmpty(name) && !name!.StartsWith('_'))
            {
                block.Bookmarks.Add(name);
            }
        }
    }

    private static List<InlineNode> RewriteInlineText(List<InlineNode> inlines, ContextIdMarker.Marker marker)
    {
        if (string.IsNullOrEmpty(marker.Symbol))
        {
            return inlines;
        }

        // The marker may be split across several runs, so locate it on the flattened
        // text and then drop the matching character range from the individual nodes.
        var flat = InlineText.FlattenWithMap(inlines, out var map);
        var match = ContextIdMarker.FindMatch(flat);
        if (match is null)
        {
            return inlines;
        }

        var (start, length) = match.Value;
        var last = start + length;

        TextInline? firstTouched = null;
        TextInline? lastTouched = null;

        for (var i = 0; i < map.Count; i++)
        {
            var (nodeStart, nodeLength, textNode) = map[i];
            var nodeEnd = nodeStart + nodeLength;

            if (nodeEnd <= start || nodeStart >= last)
            {
                continue;
            }

            var removeStart = Math.Max(start, nodeStart) - nodeStart;
            var removeEnd = Math.Min(last, nodeEnd) - nodeStart;
            textNode.Text = textNode.Text.Remove(removeStart, removeEnd - removeStart);

            firstTouched ??= textNode;
            lastTouched = textNode;
        }

        // The marker leaves a stray space behind ("Titolo {#ID_X}" -> "Titolo "); trim it.
        if (firstTouched is not null)
        {
            firstTouched.Text = firstTouched.Text.TrimStart();
        }

        if (lastTouched is not null)
        {
            lastTouched.Text = lastTouched.Text.TrimEnd();
        }

        return inlines;
    }

    // ---------------------------------------------------------------- inlines

    private List<InlineNode> BuildInlines(
        Paragraph paragraph,
        MainDocumentPart main,
        Dictionary<string, ImageResource> imageCache,
        List<ImageResource> orderedImages,
        out string? lastImage,
        out (int Width, int Height) lastImageSize)
    {
        var inlines = new List<InlineNode>();
        lastImage = null;
        lastImageSize = (0, 0);

        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    foreach (var node in BuildRunInlines(run, main, imageCache, orderedImages, out var image, out var size))
                    {
                        inlines.Add(node);
                        if (image is not null)
                        {
                            lastImage = image;
                            lastImageSize = size;
                        }
                    }

                    break;
                case Hyperlink link:
                    inlines.Add(BuildHyperlink(link, main, imageCache, orderedImages));
                    break;
                case SimpleField simpleField:
                    foreach (var node in BuildSimpleField(simpleField, main, imageCache, orderedImages))
                    {
                        inlines.Add(node);
                    }

                    break;
                case SdtRun sdtRun:
                    foreach (var run2 in sdtRun.Descendants<Run>())
                    {
                        inlines.AddRange(BuildRunInlines(run2, main, imageCache, orderedImages, out _, out _));
                    }

                    break;
                case BookmarkStart:
                case BookmarkEnd:
                case ProofError:
                    break;
            }
        }

        return inlines;
    }

    private IEnumerable<InlineNode> BuildRunInlines(
        Run run,
        MainDocumentPart main,
        Dictionary<string, ImageResource> imageCache,
        List<ImageResource> orderedImages,
        out string? imageFileName,
        out (int Width, int Height) imageSize)
    {
        imageFileName = null;
        imageSize = (0, 0);
        var format = ReadRunFormat(run);
        var nodes = new List<InlineNode>();

        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case Text text when text.Text.Length > 0:
                    nodes.Add(new TextInline
                    {
                        Text = text.Text,
                        Bold = format.Bold,
                        Italic = format.Italic,
                        Underline = format.Underline,
                        Strike = format.Strike,
                        Superscript = format.Superscript,
                        Subscript = format.Subscript,
                        StyleId = format.StyleId,
                    });
                    break;
                case Break br:
                    nodes.Add(new LineBreakInline());
                    break;
                case TabChar:
                    nodes.Add(new TextInline { Text = "\t", Bold = format.Bold, Italic = format.Italic });
                    break;
                case Drawing or Picture:
                    var (fileName, size) = ResolveImage(run, main, imageCache, orderedImages);
                    if (fileName is not null)
                    {
                        imageFileName = fileName;
                        imageSize = size;
                        nodes.Add(new TextInline
                        {
                            Text = string.Empty,
                            ImageFileName = fileName,
                            Bold = format.Bold,
                            Italic = format.Italic,
                        });
                    }

                    break;
            }
        }

        return nodes;
    }

    private HyperlinkInline BuildHyperlink(
        Hyperlink link,
        MainDocumentPart main,
        Dictionary<string, ImageResource> imageCache,
        List<ImageResource> orderedImages)
    {
        var node = new HyperlinkInline();
        var relId = link.Id?.Value;
        if (!string.IsNullOrEmpty(relId))
        {
            try
            {
                node.ExternalUrl = main.HyperlinkRelationships
                    .FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
            }
            catch (KeyNotFoundException)
            {
                node.ExternalUrl = null;
            }
        }

        if (string.IsNullOrEmpty(node.ExternalUrl))
        {
            node.BookmarkName = link.Anchor?.Value;
        }

        foreach (var child in link.ChildElements)
        {
            if (child is Run run)
            {
                node.Children.AddRange(
                    BuildRunInlines(run, main, imageCache, orderedImages, out _, out _));
            }
        }

        return node;
    }

    private IEnumerable<InlineNode> BuildSimpleField(
        SimpleField field,
        MainDocumentPart main,
        Dictionary<string, ImageResource> imageCache,
        List<ImageResource> orderedImages)
    {
        if (IsIndexField(field.Instruction?.Value))
        {
            // XE fields are invisible in the output; they are consumed as index keywords.
            return Array.Empty<InlineNode>();
        }

        var nodes = new List<InlineNode>();
        foreach (var run in field.Elements<Run>())
        {
            nodes.AddRange(BuildRunInlines(run, main, imageCache, orderedImages, out _, out _));
        }

        return nodes;
    }

    // ---------------------------------------------------------------- tables

    private TableBlock ParseTable(
        Table table,
        MainDocumentPart main,
        IReadOnlyDictionary<int, List<NumberingLevel>> numberingMap,
        IReadOnlyDictionary<string, StyleInfo> styleMap,
        Dictionary<string, ImageResource> imageCache,
        List<ImageResource> orderedImages)
    {
        var block = new TableBlock();
        var rows = table.Elements<OpenXmlTableRow>().ToList();

        foreach (var row in rows)
        {
            var tableRow = new ModelTableRow();
            foreach (var cell in row.Elements<OpenXmlTableCell>())
            {
                var tableCell = new ModelTableCell
                {
                    GridSpan = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1,
                };

                var container = new ParsedDocument();
                foreach (var element in cell.ChildElements)
                {
                    if (element is Paragraph paragraph)
                    {
                        AddParagraph(paragraph, container, main, numberingMap, styleMap, imageCache, orderedImages);
                    }
                    else if (element is Table nested)
                    {
                        container.Blocks.Add(ParseTable(nested, main, numberingMap, styleMap, imageCache, orderedImages));
                    }
                }

                tableCell.Blocks.AddRange(container.Blocks);
                tableRow.Cells.Add(tableCell);
            }

            block.Rows.Add(tableRow);
        }

        if (rows.Count > 1 && IsHeaderRow(rows[0]))
        {
            block.HasHeaderRow = true;
        }

        return block;
    }

    private static bool IsHeaderRow(OpenXmlTableRow row)
    {
        // Word stores "repeat as header row" on the row properties (w:trPr/w:tblHeader);
        // some producers only flag the individual cells, so both are honoured.
        var rowHeader = row.TableRowProperties?.Elements<OpenXmlTableHeader>().FirstOrDefault();
        if (rowHeader is not null && (rowHeader.Val is null || rowHeader.Val.Value != OnOffOnlyValues.Off))
        {
            return true;
        }

        return row.Elements<OpenXmlTableCell>().Any(IsHeaderCell);
    }

    private static bool IsHeaderCell(OpenXmlTableCell cell)
    {
        var header = cell.TableCellProperties?.Elements<OpenXmlTableHeader>().FirstOrDefault();
        return header is not null &&
               (header.Val is null || header.Val.Value != OnOffOnlyValues.Off);
    }

    // ---------------------------------------------------------------- lists

    private ListBlock? BuildList(
        Paragraph paragraph,
        IReadOnlyDictionary<int, List<NumberingLevel>> numberingMap,
        MainDocumentPart main,
        Dictionary<string, ImageResource> imageCache,
        List<ImageResource> orderedImages)
    {
        var props = paragraph.ParagraphProperties;
        var numPr = props?.NumberingProperties;
        var numId = numPr?.NumberingId?.Val?.Value;
        var ilvl = numPr?.NumberingLevelReference?.Val?.Value ?? 0;

        var ordered = true;
        if (numId.HasValue && numberingMap.TryGetValue(numId.Value, out var levels))
        {
            var level = levels.FirstOrDefault(l => l.Level == ilvl);
            ordered = level?.Ordered ?? true;
        }

        var inlines = BuildInlines(paragraph, main, imageCache, orderedImages, out _, out _);
        if (inlines.Count == 0)
        {
            return null;
        }

        var list = new ListBlock { Ordered = ordered };
        var item = new ModelListItem { Level = ilvl };
        item.Inlines.AddRange(inlines);
        list.Items.Add(item);
        return list;
    }

    private static bool IsListParagraph(ParagraphProperties? props) =>
        props?.NumberingProperties?.NumberingId?.Val?.Value is not null;

    /// <summary>
    /// A level is ordered unless it renders as a bullet or has no numbering at all.
    /// </summary>
    private static bool IsOrderedFormat(NumberFormatValues? format) =>
        format is not null &&
        format.Value != NumberFormatValues.None &&
        format.Value != NumberFormatValues.Bullet;

    // ---------------------------------------------------------------- index entries

    private static List<IndexKeyword> ExtractIndexKeywords(Paragraph paragraph)
    {
        var keywords = new List<IndexKeyword>();

        foreach (var field in paragraph.Descendants<SimpleField>())
        {
            var instruction = field.Instruction?.Value;
            if (!IsIndexField(instruction))
            {
                continue;
            }

            var keyword = ParseXeInstruction(instruction);
            if (keyword is not null)
            {
                keywords.Add(keyword);
            }
        }

        foreach (var fieldCode in paragraph.Descendants<FieldCode>())
        {
            if (!IsIndexField(fieldCode.Text))
            {
                continue;
            }

            var keyword = ParseXeInstruction(fieldCode.Text);
            if (keyword is not null)
            {
                keywords.Add(keyword);
            }
        }

        return keywords;
    }

    private static bool IsIndexField(string? instruction) =>
        !string.IsNullOrWhiteSpace(instruction) &&
        instruction!.TrimStart().StartsWith("XE", StringComparison.OrdinalIgnoreCase);

    internal static IndexKeyword? ParseXeInstruction(string? instruction)
    {
        if (string.IsNullOrWhiteSpace(instruction))
        {
            return null;
        }

        var match = XeFieldRegex.Match(instruction);
        if (!match.Success)
        {
            return null;
        }

        var main = match.Groups["main"].Value.Trim();
        if (main.Length == 0)
        {
            return null;
        }

        var sub = new[] { "sub", "sub2", "sub3" }
            .Select(name => match.Groups[name].Success ? match.Groups[name].Value.Trim().Trim('"') : null)
            .FirstOrDefault(value => !string.IsNullOrEmpty(value));

        return new IndexKeyword { Keyword = main, SubKeyword = sub };
    }

    // ---------------------------------------------------------------- styles / numbering

    private static int ResolveHeadingLevel(string? styleId, string? styleName, int outlineLevel)
    {
        // The style id is normally "Heading1"; the displayed name may be localised
        // ("Titolo 1", "Título 1", "Überschrift 1"), so both are accepted.
        foreach (var candidate in new[] { styleId, styleName })
        {
            if (string.IsNullOrEmpty(candidate))
            {
                continue;
            }

            var match = StyleHeadingRegex.Match(candidate);
            if (match.Success)
            {
                return int.Parse(match.Groups["level"].Value, CultureInfo.InvariantCulture);
            }
        }

        // Outlines 0..8 in Word map to heading levels 1..9; only levels 1..3 are used for pages/toc depth.
        return outlineLevel >= 0 && outlineLevel < 9 ? outlineLevel + 1 : 0;
    }

    private static int ResolveOutlineLevel(
        ParagraphProperties? props,
        string? styleId,
        IReadOnlyDictionary<string, StyleInfo> styleMap)
    {
        var direct = props?.OutlineLevel?.Val?.Value;
        if (direct.HasValue)
        {
            return direct.Value;
        }

        if (!string.IsNullOrEmpty(styleId) && styleMap.TryGetValue(styleId!, out var info) && info.OutlineLevel >= 0)
        {
            return info.OutlineLevel;
        }

        return -1;
    }

    private static bool IsCodeStyle(string? styleId, IReadOnlyDictionary<string, StyleInfo> styleMap)
    {
        if (string.IsNullOrEmpty(styleId))
        {
            return false;
        }

        return styleId.Contains("code", StringComparison.OrdinalIgnoreCase)
            || styleId.Contains("codice", StringComparison.OrdinalIgnoreCase)
            || styleId.Contains("sorgente", StringComparison.OrdinalIgnoreCase)
            || (styleMap.TryGetValue(styleId, out var info) &&
                info.FontName?.Contains("Consolas", StringComparison.OrdinalIgnoreCase) == true)
            || (styleMap.TryGetValue(styleId, out var info2) &&
                info2.FontName?.Contains("Courier", StringComparison.OrdinalIgnoreCase) == true);
    }

    private static (bool Bold, bool Italic, bool Underline, bool Strike, bool Superscript, bool Subscript, string? StyleId) ReadRunFormat(Run run)
    {
        var props = run.RunProperties;
        var styleId = props?.RunStyle?.Val?.Value;
        var vertical = props?.VerticalTextAlignment?.Val;

        return (
            props?.Bold is not null && props.Bold.Val?.Value != false,
            props?.Italic is not null && props.Italic.Val?.Value != false,
            props?.Underline is not null && props.Underline.Val?.Value != UnderlineValues.None,
            props?.Strike is not null && props.Strike.Val?.Value != false,
            vertical?.Value == VerticalPositionValues.Superscript,
            vertical?.Value == VerticalPositionValues.Subscript,
            styleId);
    }

    private static Dictionary<string, StyleInfo> BuildStyleMap(MainDocumentPart main)
    {
        var map = new Dictionary<string, StyleInfo>(StringComparer.Ordinal);
        var styles = main.StyleDefinitionsPart?.Styles;
        if (styles is null)
        {
            return map;
        }

        foreach (var style in styles.Elements<Style>())
        {
            var id = style.StyleId?.Value;
            if (string.IsNullOrEmpty(id))
            {
                continue;
            }

            var outline = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value ?? -1;
            var font = style.StyleRunProperties?.RunFonts?.Ascii?.Value;

            // Word localises the displayed name ("Titolo 1" in an Italian document) while
            // the style id usually stays "Heading1". Both are checked so a document
            // authored in any language still produces headings.
            map[id!] = new StyleInfo(outline, font, style.Type?.Value, style.StyleName?.Val?.Value);
        }

        return map;
    }

    private static Dictionary<int, List<NumberingLevel>> BuildNumberingMap(MainDocumentPart main)
    {
        var result = new Dictionary<int, List<NumberingLevel>>();
        var numbering = main.NumberingDefinitionsPart?.Numbering;
        if (numbering is null)
        {
            return result;
        }

        var abstractByNum = numbering.Elements<NumberingInstance>()
            .Where(n => n.NumberID?.Value is not null)
            .ToDictionary(
                n => n.NumberID!.Value,
                n => n.AbstractNumId?.Val?.Value);

        var abstracts = numbering.Elements<AbstractNum>()
            .Where(a => a.AbstractNumberId?.Value is not null)
            .ToDictionary(a => a.AbstractNumberId!.Value!, a => a);

        foreach (var (numId, abstractId) in abstractByNum)
        {
            if (abstractId is null || !abstracts.TryGetValue(abstractId.Value, out var abstractNum))
            {
                continue;
            }

            var levels = new List<NumberingLevel>();
            foreach (var level in abstractNum.Elements<Level>())
            {
                var lvl = level.LevelIndex?.Value ?? 0;
                var format = level.NumberingFormat?.Val?.Value;
                levels.Add(new NumberingLevel(lvl, IsOrderedFormat(format)));
            }

            result[numId] = levels;
        }

        return result;
    }

    private static string? ReadCoreTitle(WordprocessingDocument document)
    {
        var props = document.PackageProperties;
        return string.IsNullOrWhiteSpace(props.Title) ? null : props.Title.Trim();
    }

    private static bool ContainsPageBreak(Paragraph paragraph) =>
        paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page);

    /// <summary>
    /// True when a paragraph carries nothing but line/page breaks, i.e. it is a manual
    /// page break rather than content.
    /// </summary>
    private static bool IsBlankApartFromBreaks(List<InlineNode> inlines) =>
        inlines.All(node => node is LineBreakInline ||
                            (node is TextInline text && string.IsNullOrWhiteSpace(text.Text)));

    private static string Slugify(string value) =>
        new string(value.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();

    // ---------------------------------------------------------------- images

    private static (string? FileName, (int Width, int Height) Size) ResolveImage(
        Run run,
        MainDocumentPart main,
        Dictionary<string, ImageResource> cache,
        List<ImageResource> ordered)
    {
        var blip = run.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        var relId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(relId))
        {
            return (null, (0, 0));
        }

        if (!cache.TryGetValue(relId!, out var resource))
        {
            var part = (ImagePart)main.GetPartById(relId!);
            var fileName = Path.GetFileName(part.Uri.OriginalString);
            using var buffer = new MemoryStream();
            using (var content = part.GetStream())
            {
                content.CopyTo(buffer);
            }

            resource = new ImageResource
            {
                FileName = fileName,
                Content = buffer.ToArray(),
                ContentType = part.ContentType ?? "image/png",
            };

            cache[relId!] = resource;
            ordered.Add(resource);
        }

        var size = ReadExtent(run);
        return (resource.FileName, size);
    }

    private static (int Width, int Height) ReadExtent(Run run)
    {
        var extent = run.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
        if (extent?.Cx is null || extent.Cy is null)
        {
            return (0, 0);
        }

        // EMU: 914400 per inch; convert to CSS pixels at 96 dpi.
        const double emuPerPixel = 914400.0 / 96.0;
        return (
            (int)Math.Round(extent.Cx.Value / emuPerPixel),
            (int)Math.Round(extent.Cy.Value / emuPerPixel));
    }

    private sealed record StyleInfo(int OutlineLevel, string? FontName, StyleValues? Type, string? Name);

    private sealed record NumberingLevel(int Level, bool Ordered);
}
