using System.Text;

namespace Word2Chm.Core.Common;

/// <summary>
/// Rewrites every non-ASCII character as a numeric entity. The Help compiler reads the
/// topics as ANSI, so literal UTF-8 punctuation (typographic quotes, dashes) is
/// mis-decoded and the resulting CHM renders the wrong glyphs. Entities keep the files
/// pure ASCII while the viewer still shows the original characters.
/// </summary>
public static class AsciiEncoder
{
    public static string Encode(string text)
    {
        var builder = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            if (ch < 0x80)
            {
                builder.Append(ch);
            }
            else
            {
                builder.Append("&#").Append((int)ch).Append(';');
            }
        }

        return builder.ToString();
    }
}
