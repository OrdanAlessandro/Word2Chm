using System.Globalization;
using System.Text;

namespace Word2Chm.Core.Common;

/// <summary>
/// Produces URL/HTML-safe identifiers from arbitrary text and guarantees that
/// identifiers stay unique within a single help project.
/// </summary>
public sealed class Slugger
{
    private readonly Dictionary<string, int> _used = new(StringComparer.OrdinalIgnoreCase);

    public string Slug(string text)
    {
        var builder = new StringBuilder(text.Length);
        var pendingSeparator = false;

        foreach (var ch in Deaccent(text))
        {
            if (char.IsLetterOrDigit(ch))
            {
                if (pendingSeparator && builder.Length > 0)
                {
                    builder.Append('-');
                }

                builder.Append(char.ToLowerInvariant(ch));
                pendingSeparator = false;
            }
            else
            {
                pendingSeparator = true;
            }
        }

        var slug = builder.Length == 0 ? "section" : builder.ToString();
        return MakeUnique(slug);
    }

    /// <summary>
    /// Registers a caller-provided identifier so that later slugs cannot collide with it.
    /// </summary>
    public void Reserve(string candidate)
    {
        if (string.IsNullOrEmpty(candidate))
        {
            return;
        }

        _used[candidate] = _used.TryGetValue(candidate, out var count) ? count + 1 : 1;
    }

    private string MakeUnique(string slug)
    {
        if (!_used.TryGetValue(slug, out var count))
        {
            _used[slug] = 1;
            return slug;
        }

        string candidate;
        do
        {
            count++;
            candidate = slug + "-" + count.ToString(CultureInfo.InvariantCulture);
        }
        while (_used.ContainsKey(candidate));

        _used[slug] = count;
        _used[candidate] = 1;
        return candidate;
    }

    /// <summary>
    /// Removes diacritics so that accented titles still produce readable anchors.
    /// Uses an explicit table instead of Unicode decomposition, which is unavailable
    /// under invariant globalization.
    /// </summary>
    public static string Deaccent(string text)
    {
        var builder = new StringBuilder(text.Length);

        foreach (var ch in text)
        {
            builder.Append(AccentMap.TryGetValue(ch, out var replacement) ? replacement : ch);
        }

        return builder.ToString();
    }

    private static readonly Dictionary<char, char> AccentMap = new()
    {
        ['à'] = 'a', ['á'] = 'a', ['â'] = 'a', ['ã'] = 'a', ['ä'] = 'a', ['å'] = 'a',
        ['è'] = 'e', ['é'] = 'e', ['ê'] = 'e', ['ë'] = 'e',
        ['ì'] = 'i', ['í'] = 'i', ['î'] = 'i', ['ï'] = 'i',
        ['ò'] = 'o', ['ó'] = 'o', ['ô'] = 'o', ['õ'] = 'o', ['ö'] = 'o',
        ['ù'] = 'u', ['ú'] = 'u', ['û'] = 'u', ['ü'] = 'u',
        ['ý'] = 'y', ['ÿ'] = 'y', ['ñ'] = 'n', ['ç'] = 'c',
        ['À'] = 'A', ['Á'] = 'A', ['Â'] = 'A', ['Ã'] = 'A', ['Ä'] = 'A', ['Å'] = 'A',
        ['È'] = 'E', ['É'] = 'E', ['Ê'] = 'E', ['Ë'] = 'E',
        ['Ì'] = 'I', ['Í'] = 'I', ['Î'] = 'I', ['Ï'] = 'I',
        ['Ò'] = 'O', ['Ó'] = 'O', ['Ô'] = 'O', ['Õ'] = 'O', ['Ö'] = 'O',
        ['Ù'] = 'U', ['Ú'] = 'U', ['Û'] = 'U', ['Ü'] = 'U',
        ['Ý'] = 'Y', ['Ñ'] = 'N', ['Ç'] = 'C',
    };
}
