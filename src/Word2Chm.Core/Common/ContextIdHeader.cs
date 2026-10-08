using System.Globalization;
using System.Text.RegularExpressions;

namespace Word2Chm.Core.Common;

/// <summary>
/// Reads the C++ header that supplies the numeric context IDs. The header is the single
/// source of truth: the document refers to a symbol as <c>{#IDH_AXES}</c> and the number
/// comes from the matching <c>#define</c>, so the same header can back several language
/// editions and keep the IDs identical across all of them.
/// </summary>
public sealed partial class ContextIdHeader
{
    /// <summary>
    /// Matches a numeric <c>#define</c> at the start of a line. The guarded include of the
    /// header (<c>#ifndef/#define _X_H_</c>) has no number after the symbol and is therefore
    /// ignored, as are macros that expand to another macro.
    /// </summary>
    [GeneratedRegex(
        @"^[ \t]*#[ \t]*define[ \t]+(?<sym>[A-Za-z_][A-Za-z0-9_]*)[ \t]+\(?[ \t]*(?<num>0[xX][0-9A-Fa-f]+|[0-9]+)[uUlL]*[ \t]*\)?",
        RegexOptions.Multiline)]
    private static partial Regex DefineRegex();

    /// <summary>Removes <c>//</c> comments so an example inside a comment is not read as a definition.</summary>
    [GeneratedRegex(@"//[^\n]*")]
    private static partial Regex LineCommentRegex();

    private readonly Dictionary<string, int> _definitions = new(StringComparer.Ordinal);

    /// <summary>Symbol name to numeric context ID, as read from the header.</summary>
    public IReadOnlyDictionary<string, int> Definitions => _definitions;

    public static ContextIdHeader Parse(string text)
    {
        var header = new ContextIdHeader();
        var source = LineCommentRegex().Replace(text, string.Empty);

        foreach (Match match in DefineRegex().Matches(source))
        {
            var symbol = match.Groups["sym"].Value;
            var raw = match.Groups["num"].Value;
            var value = raw.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                ? int.Parse(raw[2..], NumberStyles.HexNumber, CultureInfo.InvariantCulture)
                : int.Parse(raw, CultureInfo.InvariantCulture);

            // Redefining a symbol is legal in C and the last value is the effective one.
            header._definitions[symbol] = value;
        }

        return header;
    }

    public static ContextIdHeader Load(string path) => Parse(File.ReadAllText(path));

    public bool TryGetId(string symbol, out int id) => _definitions.TryGetValue(symbol, out id);
}
