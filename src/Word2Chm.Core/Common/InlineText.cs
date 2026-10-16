using System.Text;
using Word2Chm.Core.Model;

namespace Word2Chm.Core.Common;

/// <summary>Maps a character position in the flattened text back to a text node.</summary>
public sealed record TextSpan(int Start, int Length, TextInline Node);

/// <summary>Flattens inline nodes back to plain text.</summary>
public static class InlineText
{
    public static string Flatten(IEnumerable<InlineNode> nodes)
    {
        var builder = new StringBuilder();
        Append(nodes, builder);
        return builder.ToString().Trim();
    }

    /// <summary>
    /// Flattens the text nodes and records where each one starts, so callers can
    /// edit the original nodes based on a position found in the flattened string.
    /// </summary>
    public static string FlattenWithMap(IEnumerable<InlineNode> nodes, out List<TextSpan> map)
    {
        map = new List<TextSpan>();
        var builder = new StringBuilder();
        AppendMapped(nodes, builder, map);
        return builder.ToString();
    }

    private static void Append(IEnumerable<InlineNode> nodes, StringBuilder builder)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case TextInline text:
                    builder.Append(text.Text);
                    break;
                case HyperlinkInline link:
                    Append(link.Children, builder);
                    break;
                case LineBreakInline:
                    builder.Append(' ');
                    break;
            }
        }
    }

    private static void AppendMapped(IEnumerable<InlineNode> nodes, StringBuilder builder, List<TextSpan> map)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case TextInline { ImageFileName: null } text:
                    map.Add(new TextSpan(builder.Length, text.Text.Length, text));
                    builder.Append(text.Text);
                    break;
                case HyperlinkInline link:
                    AppendMapped(link.Children, builder, map);
                    break;
                case LineBreakInline:
                    builder.Append(' ');
                    break;
            }
        }
    }
}
