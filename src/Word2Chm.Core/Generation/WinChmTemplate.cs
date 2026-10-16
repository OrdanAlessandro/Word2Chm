using System.Globalization;
using System.Net;
using System.Text;
using Word2Chm.Core.Common;
using Word2Chm.Core.Model;

namespace Word2Chm.Core.Generation;

/// <summary>
/// Wraps rendered page content in the WinCHM "fixedtop" skin: a fixed header carrying the
/// breadcrumb, the page title and the previous/next buttons, plus a scrolling content area
/// with a footer.
/// </summary>
public static class WinChmTemplate
{
    public const string TemplateFileName = "fixedtop.htm";
    public const string StyleFileName = "winchm_template_style.css";
    public const string ScriptFileName = "winchm_template_script.js";

    /// <summary>
    /// Binary assets of the skin that must be copied next to the pages and declared in the
    /// project file, otherwise the buttons do not render inside the CHM.
    /// </summary>
    public static readonly string[] ImageFileNames =
    {
        "btn_prev_n.gif",
        "btn_prev_g.gif",
        "btn_next_n.gif",
        "btn_next_g.gif",
    };

    /// <summary>
    /// Applies the skin to already-rendered page content. The placeholders are replaced
    /// literally because the skin is written by hand with spaces inside the markers.
    /// </summary>
    public static string Apply(
        string template,
        string title,
        string content,
        string navigation,
        string footer,
        string? previousFile,
        string? nextFile,
        string projectCssFileName)
    {
        var result = template
            .Replace("($title$)", WebUtility.HtmlEncode(title), StringComparison.Ordinal)
            .Replace("($content$)", content, StringComparison.Ordinal)
            .Replace("($navigation$)", navigation, StringComparison.Ordinal)
            .Replace("($footer$)", WebUtility.HtmlEncode(footer), StringComparison.Ordinal);

        result = result
            .Replace("<img src=\"btn_prev_n.gif\">", ButtonImage("btn_prev_n.gif", previousFile), StringComparison.Ordinal)
            .Replace("<img src=\"btn_next_n.gif\">", ButtonImage("btn_next_n.gif", nextFile), StringComparison.Ordinal);

        // The skin only links its own stylesheet; the project stylesheet is added as well so
        // tables, code blocks and list levels keep their formatting.
        var link = $"<link rel=\"stylesheet\" type=\"text/css\" href=\"{projectCssFileName}\">";
        result = result.Replace("</head>", link + Environment.NewLine + "</head>", StringComparison.OrdinalIgnoreCase);

        return AsciiEncoder.Encode(result);
    }

    /// <summary>
    /// Builds the breadcrumb shown above the page title. Each ancestor links to the anchor
    /// generated for its heading, so the trail is navigable.
    /// </summary>
    public static string BuildNavigation(IEnumerable<(string Title, string? Local)> ancestors)
    {
        var parts = new List<string>();
        foreach (var (title, local) in ancestors)
        {
            var text = WebUtility.HtmlEncode(title);
            parts.Add(string.IsNullOrEmpty(local)
                ? text
                : $"<a href=\"{WebUtility.HtmlEncode(local)}\">{text}</a>");
        }

        return string.Join(" &gt; ", parts);
    }

    /// <summary>Pointing the whole button at the neighbouring page keeps the image clickable.</summary>
    private static string ButtonImage(string fileName, string? target) =>
        string.IsNullOrEmpty(target)
            ? $"<img src=\"{fileName}\">"
            : $"<a href=\"{WebUtility.HtmlEncode(target)}\"><img src=\"{fileName}\"></a>";

    /// <summary>Reads the skin, stripping the UTF-8 BOM that editors sometimes leave behind.</summary>
    public static string ReadTemplate(string path) =>
        File.ReadAllText(path, Encoding.UTF8).TrimStart('\uFEFF');
}
