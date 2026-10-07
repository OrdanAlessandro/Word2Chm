using System.Globalization;

namespace Word2Chm.Core.Generation;

/// <summary>Embedded stylesheet shared by every generated page.</summary>
public static class CssGenerator
{
    public const string FileName = "help.css";

    /// <summary>
    /// Placeholder standing in for the configured body size, so the stylesheet stays a
    /// readable raw string instead of an interpolated one where every brace would need
    /// doubling.
    /// </summary>
    private const string BodyFontSizeToken = "__BODY_FONT_SIZE__";

    public static string Generate(double bodyFontSizePt = BuildOptions.DefaultBodyFontSizePt) =>
        Template.Replace(
            BodyFontSizeToken,
            bodyFontSizePt.ToString("0.##", CultureInfo.InvariantCulture),
            StringComparison.Ordinal);

    private static string Template => """
        :root {
            color-scheme: light dark;
            --fg: #1f2328;
            --bg: #ffffff;
            --muted: #57606a;
            --border: #d0d7de;
            --accent: #0969da;
            --code-bg: #f6f8fa;
            --body-font-size: __BODY_FONT_SIZE__pt;
        }

        * { box-sizing: border-box; }

        body {
            margin: 0;
            padding: 0;
            font-family: "Segoe UI", Tahoma, Arial, sans-serif;
            font-size: var(--body-font-size);
            line-height: 1.55;
            color: var(--fg);
            background: var(--bg);
        }

        /* The WinCHM skin sizes its content div through an id selector, which outranks the
           body rule above. The size is repeated here so the configured value also applies
           inside the skin. */
        #winchm_template_content { font-size: var(--body-font-size); }

        .page {
            max-width: 900px;
            margin: 0 auto;
            padding: 28px 32px 48px;
        }

        h1, h2, h3, h4, h5, h6 {
            line-height: 1.25;
            margin: 1.4em 0 0.6em;
            font-weight: 600;
        }

        h1 { font-size: 1.9em; border-bottom: 2px solid var(--border); padding-bottom: 0.3em; }
        h2 { font-size: 1.5em; }
        h3 { font-size: 1.25em; }
        h4 { font-size: 1.1em; }

        p { margin: 0.7em 0; }

        a { color: var(--accent); text-decoration: none; }
        a:hover { text-decoration: underline; }

        ul, ol { margin: 0.7em 0; padding-left: 1.8em; }
        li { margin: 0.25em 0; }
        li.level-1 { margin-left: 1.2em; }
        li.level-2 { margin-left: 2.4em; }
        li.level-3 { margin-left: 3.6em; }

        table {
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
        }

        th, td {
            /* Word's "Grigliatabella" style draws 0.5pt black lines; the shared --border
               token is a light grey, which made every table look washed out. */
            border: 1px solid #000000;
            padding: 6px 10px;
            text-align: left;
            vertical-align: top;
        }

        th { background: var(--code-bg); font-weight: 600; }

        pre.code {
            background: var(--code-bg);
            border: 1px solid var(--border);
            border-radius: 6px;
            padding: 12px;
            overflow-x: auto;
            font-family: Consolas, "Courier New", monospace;
            font-size: 0.9em;
        }

        code {
            font-family: Consolas, "Courier New", monospace;
            background: var(--code-bg);
            padding: 0.1em 0.35em;
            border-radius: 4px;
        }

        figure { margin: 1em 0; text-align: center; }
        figure img { max-width: 100%; height: auto; }
        figcaption { margin-top: 0.4em; font-size: 0.9em; color: var(--muted); }

        hr.pagebreak { border: none; border-top: 1px solid var(--border); margin: 1.5em 0; }

        ul.toc { list-style: none; padding-left: 0; }
        ul.toc ul { list-style: none; padding-left: 1.4em; }
        ul.toc a { font-weight: 500; }

        .footer {
            max-width: 900px;
            margin: 0 auto;
            padding: 12px 32px 24px;
            border-top: 1px solid var(--border);
            color: var(--muted);
            font-size: 0.85em;
        }

        @media (prefers-color-scheme: dark) {
            :root {
                --fg: #e6edf3;
                --bg: #0d1117;
                --muted: #8b949e;
                --border: #30363d;
                --accent: #58a6ff;
                --code-bg: #161b22;
            }
        }
        """;
}
