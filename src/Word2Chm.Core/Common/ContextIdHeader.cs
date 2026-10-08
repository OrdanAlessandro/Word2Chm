using System.Globalization;
using System.Text.RegularExpressions;

namespace Word2Chm.Core.Common;

/// <summary>
/// Reads the C++ header that supplies the numeric context IDs. The header is the single
/// source of truth: the document refers to a symbol as <c>{#IDH_AXES}</c> and the number is
/// read from the matching <c>#define</c> or <c>enum</c>, so the same header can back several
/// language editions and keep the IDs identical across all of them.
/// </summary>
public sealed partial class ContextIdHeader
{
    /// <summary>
    /// Matches, at the start of a line, either a numeric <c>#define</c> or a whole
    /// <c>enum</c> block. The guarded include of the header (<c>#ifndef/#define _X_H_</c>) has
    /// no number after the symbol and is therefore ignored, as are macros that expand to
    /// another macro.
    /// </summary>
    [GeneratedRegex(
        @"(?<define>^[ \t]*#[ \t]*define[ \t]+(?<sym>[A-Za-z_][A-Za-z0-9_]*)[ \t]+\(?[ \t]*(?<num>0[xX][0-9A-Fa-f]+|[0-9]+)[uUlL]*[ \t]*\)?)" +
        @"|(?<enum>\benum\b[^{;]*\{(?<ebody>[^{}]*)\})",
        RegexOptions.Multiline)]
    private static partial Regex DefinitionRegex();

    /// <summary>Removes <c>//</c> comments so an example inside a comment is not read as a definition.</summary>
    [GeneratedRegex(@"//[^\n]*")]
    private static partial Regex LineCommentRegex();

    /// <summary>Removes <c>/* ... */</c> comments, which can span several lines.</summary>
    [GeneratedRegex(@"/\*.*?\*/", RegexOptions.Singleline)]
    private static partial Regex BlockCommentRegex();

    /// <summary>One enumerator: the name, followed by an optional <c>= expression</c>.</summary>
    [GeneratedRegex(@"^(?<sym>[A-Za-z_][A-Za-z0-9_]*)[ \t]*(?:=[ \t]*(?<expr>.*))?$", RegexOptions.Singleline)]
    private static partial Regex EnumeratorRegex();

    private readonly Dictionary<string, int> _definitions = new(StringComparer.Ordinal);

    /// <summary>Symbol name to numeric context ID, as read from the header.</summary>
    public IReadOnlyDictionary<string, int> Definitions => _definitions;

    public static ContextIdHeader Parse(string text)
    {
        var header = new ContextIdHeader();
        var source = StripComments(text);

        // Matches are returned left to right, so a symbol defined by a #define and then by an
        // enumerator (or the other way round) keeps the last value seen, which is what the C
        // compiler does too.
        foreach (Match match in DefinitionRegex().Matches(source))
        {
            if (match.Groups["define"].Success)
            {
                header._definitions[match.Groups["sym"].Value] = ParseIntegerLiteral(match.Groups["num"].Value);
            }
            else
            {
                header.ReadEnum(match.Groups["ebody"].Value);
            }
        }

        return header;
    }

    /// <summary>
    /// Reads the body of an <c>enum</c>. An enumerator without <c>=</c> takes the previous
    /// value plus one (starting from 0). An enumerator whose expression cannot be resolved
    /// from the symbols already known is skipped rather than guessed, so it simply shows up as
    /// a missing ID; the run of implicit values that follows it is unknown as well.
    /// </summary>
    private void ReadEnum(string body)
    {
        int? previous = null;

        foreach (var raw in body.Split(','))
        {
            var item = raw.Trim();
            if (item.Length == 0)
            {
                continue;
            }

            var match = EnumeratorRegex().Match(item);
            if (!match.Success)
            {
                previous = null;
                continue;
            }

            var symbol = match.Groups["sym"].Value;
            if (!match.Groups["expr"].Success)
            {
                if (previous is null)
                {
                    continue;
                }

                previous += 1;
                _definitions[symbol] = previous.Value;
                continue;
            }

            if (ExpressionEvaluator.TryEvaluate(match.Groups["expr"].Value.Trim(), _definitions, out var value))
            {
                previous = value;
                _definitions[symbol] = value;
            }
            else
            {
                previous = null;
            }
        }
    }

    private static string StripComments(string text) =>
        LineCommentRegex().Replace(BlockCommentRegex().Replace(text, string.Empty), string.Empty);

    private static int ParseIntegerLiteral(string raw) =>
        raw.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
            ? int.Parse(raw[2..], NumberStyles.HexNumber, CultureInfo.InvariantCulture)
            : int.Parse(raw, CultureInfo.InvariantCulture);

    public static ContextIdHeader Load(string path) => Parse(File.ReadAllText(path));

    public bool TryGetId(string symbol, out int id) => _definitions.TryGetValue(symbol, out id);

    /// <summary>
    /// Evaluates the constant expression that follows <c>=</c> in an enumerator. It covers
    /// integer literals, references to symbols already defined and the usual integer
    /// operators. Anything else (casts, function calls, symbols defined later in the file)
    /// reports failure instead of producing a wrong ID.
    /// </summary>
    private static class ExpressionEvaluator
    {
        private enum Kind
        {
            Number,
            Symbol,
            Operator,
            LeftParen,
            RightParen,
        }

        private readonly record struct Token(Kind Kind, long Number, string Text);

        public static bool TryEvaluate(string expression, IReadOnlyDictionary<string, int> known, out int value)
        {
            value = 0;

            if (!Tokenize(expression, out var tokens))
            {
                return false;
            }

            var position = 0;
            if (!ParseBinary(tokens, ref position, 0, known, out var result) || position != tokens.Count)
            {
                return false;
            }

            // A value outside int range is a real error in the header, not a reason to skip it.
            value = checked((int)result);
            return true;
        }

        private static bool Tokenize(string text, out List<Token> tokens)
        {
            tokens = new List<Token>();
            var index = 0;

            while (index < text.Length)
            {
                var current = text[index];
                if (char.IsWhiteSpace(current))
                {
                    index++;
                    continue;
                }

                if (char.IsDigit(current))
                {
                    var start = index;
                    var hexadecimal = current == '0' && index + 1 < text.Length && text[index + 1] is 'x' or 'X';
                    var binary = current == '0' && index + 1 < text.Length && text[index + 1] is 'b' or 'B';

                    if (hexadecimal || binary)
                    {
                        index += 2;
                    }
                    else
                    {
                        index++;
                    }

                    while (index < text.Length && (hexadecimal
                               ? Uri.IsHexDigit(text[index])
                               : binary ? text[index] is '0' or '1' : char.IsDigit(text[index])))
                    {
                        index++;
                    }

                    var raw = text[start..index];
                    while (index < text.Length && text[index] is 'u' or 'U' or 'l' or 'L')
                    {
                        index++;
                    }

                    if (!TryParseNumber(raw, out var number))
                    {
                        return false;
                    }

                    tokens.Add(new Token(Kind.Number, number, raw));
                    continue;
                }

                if (char.IsLetter(current) || current == '_')
                {
                    var start = index++;
                    while (index < text.Length && (char.IsLetterOrDigit(text[index]) || text[index] == '_'))
                    {
                        index++;
                    }

                    tokens.Add(new Token(Kind.Symbol, 0, text[start..index]));
                    continue;
                }

                if (index + 1 < text.Length && text[index..(index + 2)] is "<<" or ">>")
                {
                    tokens.Add(new Token(Kind.Operator, 0, text[index..(index + 2)]));
                    index += 2;
                    continue;
                }

                switch (current)
                {
                    case '(':
                        tokens.Add(new Token(Kind.LeftParen, 0, "("));
                        break;
                    case ')':
                        tokens.Add(new Token(Kind.RightParen, 0, ")"));
                        break;
                    case '|' or '&' or '^' or '+' or '-' or '*' or '/' or '%' or '~':
                        tokens.Add(new Token(Kind.Operator, 0, current.ToString()));
                        break;
                    default:
                        return false;
                }

                index++;
            }

            return true;
        }

        private static bool TryParseNumber(string raw, out long value)
        {
            value = 0;
            try
            {
                if (raw.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                {
                    value = long.Parse(raw[2..], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                }
                else if (raw.StartsWith("0b", StringComparison.OrdinalIgnoreCase))
                {
                    value = Convert.ToInt64(raw[2..], 2);
                }
                else if (raw.Length > 1 && raw[0] == '0')
                {
                    value = Convert.ToInt64(raw[1..], 8);
                }
                else
                {
                    value = long.Parse(raw, CultureInfo.InvariantCulture);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>Precedence climbing: binds tighter than <paramref name="minimum"/> only.</summary>
        private static bool ParseBinary(
            List<Token> tokens,
            ref int position,
            int minimum,
            IReadOnlyDictionary<string, int> known,
            out long value)
        {
            if (!ParseUnary(tokens, ref position, known, out value))
            {
                return false;
            }

            while (Peek(tokens, position) is { Kind: Kind.Operator } op && Precedence(op.Text) >= minimum)
            {
                position++;
                if (!ParseBinary(tokens, ref position, Precedence(op.Text) + 1, known, out var right))
                {
                    return false;
                }

                if (!Apply(op.Text, value, right, out value))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool ParseUnary(
            List<Token> tokens,
            ref int position,
            IReadOnlyDictionary<string, int> known,
            out long value)
        {
            if (Peek(tokens, position) is { Kind: Kind.Operator } op && op.Text is "+" or "-" or "~")
            {
                position++;
                if (!ParseUnary(tokens, ref position, known, out value))
                {
                    return false;
                }

                value = op.Text switch
                {
                    "-" => -value,
                    "~" => ~value,
                    _ => value,
                };

                return true;
            }

            return ParsePrimary(tokens, ref position, known, out value);
        }

        private static bool ParsePrimary(
            List<Token> tokens,
            ref int position,
            IReadOnlyDictionary<string, int> known,
            out long value)
        {
            value = 0;
            if (position >= tokens.Count)
            {
                return false;
            }

            var token = tokens[position];
            switch (token.Kind)
            {
                case Kind.Number:
                    position++;
                    value = token.Number;
                    return true;

                case Kind.Symbol:
                    position++;
                    return known.TryGetValue(token.Text, out var id) && (value = id) == id;

                case Kind.LeftParen:
                    position++;
                    if (!ParseBinary(tokens, ref position, 0, known, out value))
                    {
                        return false;
                    }

                    if (Peek(tokens, position) is not { Kind: Kind.RightParen })
                    {
                        return false;
                    }

                    position++;
                    return true;

                default:
                    return false;
            }
        }

        private static bool Apply(string op, long left, long right, out long result)
        {
            result = 0;
            switch (op)
            {
                case "+": result = left + right; break;
                case "-": result = left - right; break;
                case "*": result = left * right; break;
                case "/":
                    if (right == 0) return false;
                    result = left / right;
                    break;
                case "%":
                    if (right == 0) return false;
                    result = left % right;
                    break;
                case "<<": result = left << (int)right; break;
                case ">>": result = left >> (int)right; break;
                case "&": result = left & right; break;
                case "^": result = left ^ right; break;
                case "|": result = left | right; break;
                default: return false;
            }

            return true;
        }

        private static int Precedence(string op) => op switch
        {
            "|" => 1,
            "^" => 2,
            "&" => 3,
            "<<" or ">>" => 4,
            "+" or "-" => 5,
            "*" or "/" or "%" => 6,
            _ => 0,
        };

        private static Token? Peek(List<Token> tokens, int position) =>
            position >= tokens.Count ? null : tokens[position];
    }
}
