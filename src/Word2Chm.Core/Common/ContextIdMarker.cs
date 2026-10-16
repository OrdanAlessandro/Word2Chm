using System.Text.RegularExpressions;

namespace Word2Chm.Core.Common;

/// <summary>
/// Reads and removes the context ID marker a user writes in a heading, e.g.
/// <c>Installazione {#IDH_INSTALLAZIONE}</c>. The marker is what keeps identifiers stable
/// across reordering and translation; the numeric value always comes from the supplied
/// header, so an older <c>{#IDH_X=123}</c> is accepted but its number is ignored.
/// </summary>
public static partial class ContextIdMarker
{
    [GeneratedRegex(@"\{\s*#\s*(?<sym>[A-Za-z_][A-Za-z0-9_]*)\s*(?:=\s*[0-9]+\s*)?\}")]
    private static partial Regex MarkerRegex();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();

    public sealed record Marker(string Symbol, string CleanText);

    /// <summary>
    /// Locates the first marker in <paramref name="text"/> and returns its start index
    /// and length, so callers can remove the exact character range.
    /// </summary>
    public static (int Start, int Length)? FindMatch(string text)
    {
        var match = MarkerRegex().Match(text);
        return match.Success ? (match.Index, match.Length) : null;
    }

    /// <summary>
    /// Extracts the marker from <paramref name="text"/>; when no marker is present the
    /// original text is returned unchanged.
    /// </summary>
    public static Marker Parse(string text)
    {
        var match = MarkerRegex().Match(text);
        if (!match.Success)
        {
            return new Marker(string.Empty, text.Trim());
        }

        var clean = CollapseWhitespace(MarkerRegex().Replace(text, " "));
        return new Marker(match.Groups["sym"].Value, clean);
    }

    private static string CollapseWhitespace(string value) =>
        WhitespaceRegex().Replace(value, " ").Trim();
}
