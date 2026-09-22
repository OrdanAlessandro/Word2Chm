using System.Net;
using System.Text;
using Word2Chm.Core.Model;

namespace Word2Chm.Core.Generation;

/// <summary>
/// Renders a <see cref="HelpPage"/> to a standalone HTML document. Navigation uses
/// relative links so the CHM viewer can resolve them without a web server.
/// </summary>
public sealed class HtmlGenerator
{
    private const string CssFileName = "help.css";

    public string GeneratePage(HelpPage page, HelpDocument document, string cssPath)
    {
        var builder = new StringBuilder();
        builder.AppendLine("<!DOCTYPE html>");
        builder.AppendLine("<html lang=\"" + WebUtility.HtmlEncode(document.Language) + "\">");
        builder.AppendLine("<head>");
        builder.AppendLine("<meta charset=\"utf-8\">");
        builder.AppendLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">");
        builder.AppendLine("<title>" + WebUtility.HtmlEncode(page.Title) + "</title>");
        builder.AppendLine("<link rel=\"stylesheet\" type=\"text/css\" href=\"" + CssFileName + "\">");
        builder.AppendLine("</head>");
        builder.AppendLine("<body>");
        builder.AppendLine("<div class=\"page\">");

        foreach (var block in page.Blocks)
        {
            RenderBlock(builder, block, document, page);
        }

        builder.AppendLine("</div>");
        builder.AppendLine("<div class=\"footer\">" + WebUtility.HtmlEncode(document.Title) + "</div>");
        builder.AppendLine("</body>");
        builder.AppendLine("</html>");
        return builder.ToString();
    }

    public string GenerateIndexPage(HelpDocument document)
    {
        var builder = new StringBuilder();
        builder.AppendLine("<!DOCTYPE html>");
        builder.AppendLine("<html lang=\"" + WebUtility.HtmlEncode(document.Language) + "\">");
        builder.AppendLine("<head>");
        builder.AppendLine("<meta charset=\"utf-8\">");
        builder.AppendLine("<title>" + WebUtility.HtmlEncode(document.Title) + "</title>");
        builder.AppendLine("<link rel=\"stylesheet\" type=\"text/css\" href=\"" + CssFileName + "\">");
        builder.AppendLine("</head>");
        builder.AppendLine("<body>");
        builder.AppendLine("<h1>" + WebUtility.HtmlEncode(document.Title) + "</h1>");
        builder.AppendLine("<ul class=\"toc\">");
        foreach (var node in document.Toc)
        {
            RenderTocNode(builder, node);
        }

        builder.AppendLine("</ul>");
        builder.AppendLine("</body>");
        builder.AppendLine("</html>");
        return builder.ToString();
    }

    private static void RenderTocNode(StringBuilder builder, TocNode node)
    {
        builder.Append("<li>");
        if (!string.IsNullOrEmpty(node.Local))
        {
            builder.Append("<a href=\"").Append(EscapeAttribute(node.Local)).Append("\">")
                .Append(WebUtility.HtmlEncode(node.Title)).Append("</a>");
        }
        else
        {
            builder.Append(WebUtility.HtmlEncode(node.Title));
        }

        if (node.Children.Count > 0)
        {
            builder.AppendLine("<ul>");
            foreach (var child in node.Children)
            {
                RenderTocNode(builder, child);
            }

            builder.AppendLine("</ul>");
        }

        builder.AppendLine("</li>");
    }

    private void RenderBlock(StringBuilder builder, DocumentBlock block, HelpDocument document, HelpPage page)
    {
        switch (block)
        {
            case HeadingBlock heading:
                RenderHeading(builder, heading);
                break;
            case ParagraphBlock paragraph:
                RenderParagraph(builder, paragraph, document, page);
                break;
            case ListBlock list:
                RenderList(builder, list, document, page);
                break;
            case TableBlock table:
                RenderTable(builder, table, document, page);
                break;
            case CodeBlock code:
                builder.Append("<pre class=\"code\">").Append(WebUtility.HtmlEncode(code.Text)).AppendLine("</pre>");
                break;
            case ImageBlock image:
                RenderImage(builder, image.FileName, image.Caption, image.Width, image.Height);
                break;
            case PageBreakBlock:
                builder.AppendLine("<hr class=\"pagebreak\">");
                break;
        }
    }

    private void RenderHeading(StringBuilder builder, HeadingBlock heading)
    {
        var level = Math.Clamp(heading.Level, 1, 6);
        builder.Append("<h").Append(level);
        if (!string.IsNullOrEmpty(heading.Anchor))
        {
            builder.Append(" id=\"").Append(EscapeAttribute(heading.Anchor)).Append('"');
        }

        builder.Append('>');
        RenderInlines(builder, heading.Inlines);
        builder.Append("</h").Append(level).AppendLine(">");
    }

    private void RenderParagraph(StringBuilder builder, ParagraphBlock paragraph, HelpDocument document, HelpPage page)
    {
        if (!string.IsNullOrEmpty(paragraph.ImageFileName) && paragraph.Inlines.Count <= 1)
        {
            RenderImage(builder, paragraph.ImageFileName!, null, 0, 0);
            return;
        }

        builder.Append("<p");
        if (!string.IsNullOrEmpty(paragraph.Anchor))
        {
            builder.Append(" id=\"").Append(EscapeAttribute(paragraph.Anchor)).Append('"');
        }

        builder.Append('>');
        RenderInlines(builder, paragraph.Inlines, document, page);
        builder.AppendLine("</p>");
    }

    private void RenderList(StringBuilder builder, ListBlock list, HelpDocument document, HelpPage page)
    {
        var tag = list.Ordered ? "ol" : "ul";
        builder.Append('<').Append(tag).AppendLine(">");

        foreach (var item in list.Items)
        {
            builder.Append("<li class=\"level-").Append(item.Level).Append("\">");
            RenderInlines(builder, item.Inlines, document, page);
            builder.AppendLine("</li>");
        }

        builder.Append("</").Append(tag).AppendLine(">");
    }

    private void RenderTable(StringBuilder builder, TableBlock table, HelpDocument document, HelpPage page)
    {
        builder.AppendLine("<table>");

        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var row = table.Rows[rowIndex];
            builder.AppendLine("<tr>");

            foreach (var cell in row.Cells)
            {
                var isHeader = table.HasHeaderRow && rowIndex == 0;
                var tag = isHeader ? "th" : "td";
                builder.Append('<').Append(tag);
                if (cell.GridSpan > 1)
                {
                    builder.Append(" colspan=\"").Append(cell.GridSpan).Append('"');
                }

                builder.Append('>');

                foreach (var block in cell.Blocks)
                {
                    RenderBlock(builder, block, document, page);
                }

                builder.Append("</").Append(tag).AppendLine(">");
            }

            builder.AppendLine("</tr>");
        }

        builder.AppendLine("</table>");
    }

    private void RenderImage(StringBuilder builder, string fileName, string? caption, int width, int height)
    {
        builder.Append("<figure><img src=\"assets/").Append(EscapeAttribute(fileName)).Append('"');
        if (width > 0)
        {
            builder.Append(" width=\"").Append(width).Append('"');
        }

        if (height > 0)
        {
            builder.Append(" height=\"").Append(height).Append('"');
        }

        builder.Append(" alt=\"").Append(EscapeAttribute(caption ?? string.Empty)).Append("\"/>");
        if (!string.IsNullOrEmpty(caption))
        {
            builder.Append("<figcaption>").Append(WebUtility.HtmlEncode(caption)).Append("</figcaption>");
        }

        builder.AppendLine("</figure>");
    }

    private void RenderInlines(StringBuilder builder, IEnumerable<InlineNode> nodes, HelpDocument? document = null, HelpPage? page = null)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case TextInline text:
                    RenderText(builder, text);
                    break;
                case HyperlinkInline link:
                    RenderHyperlink(builder, link, document, page);
                    break;
                case LineBreakInline:
                    builder.Append("<br>");
                    break;
            }
        }
    }

    private static void RenderText(StringBuilder builder, TextInline text)
    {
        if (text.ImageFileName is not null)
        {
            builder.Append("<img src=\"assets/").Append(EscapeAttribute(text.ImageFileName)).Append("\" alt=\"\"/>");
            return;
        }

        var encoded = WebUtility.HtmlEncode(text.Text);
        if (text.Bold)
        {
            encoded = "<strong>" + encoded + "</strong>";
        }

        if (text.Italic)
        {
            encoded = "<em>" + encoded + "</em>";
        }

        if (text.Underline)
        {
            encoded = "<u>" + encoded + "</u>";
        }

        if (text.Strike)
        {
            encoded = "<s>" + encoded + "</s>";
        }

        if (text.Superscript)
        {
            encoded = "<sup>" + encoded + "</sup>";
        }

        if (text.Subscript)
        {
            encoded = "<sub>" + encoded + "</sub>";
        }

        builder.Append(encoded);
    }

    private void RenderHyperlink(StringBuilder builder, HyperlinkInline link, HelpDocument? document, HelpPage? page)
    {
        string href;
        if (!string.IsNullOrEmpty(link.ExternalUrl))
        {
            href = link.ExternalUrl!;
        }
        else if (!string.IsNullOrEmpty(link.BookmarkName))
        {
            href = ResolveBookmark(link.BookmarkName!, document, page);
        }
        else
        {
            href = "#";
        }

        builder.Append("<a href=\"").Append(EscapeAttribute(href)).Append("\">");
        RenderInlines(builder, link.Children, document, page);
        builder.Append("</a>");
    }

    private static string ResolveBookmark(string bookmarkName, HelpDocument? document, HelpPage? page)
    {
        if (document is null || !document.BookmarkTargets.TryGetValue(bookmarkName, out var target))
        {
            return "#";
        }

        var prefix = page is not null && page.FileName == target.PageFileName ? string.Empty : target.PageFileName;
        return prefix + "#" + target.Anchor;
    }

    private static string EscapeAttribute(string value) => WebUtility.HtmlEncode(value);
}
