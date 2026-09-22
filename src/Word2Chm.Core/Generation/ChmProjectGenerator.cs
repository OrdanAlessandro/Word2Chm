using System.Globalization;
using System.Text;
using Word2Chm.Core.Model;

namespace Word2Chm.Core.Generation;

/// <summary>Names and identifiers used across the generated CHM project files.</summary>
public sealed class ChmProjectNames
{
    public required string BaseName { get; init; }
    public string ChmFile => BaseName + ".chm";
    public string HhpFile => BaseName + ".hhp";
    public string HhcFile => BaseName + ".hhc";
    public string HhkFile => BaseName + ".hhk";
    public string HeaderFile => BaseName + ".h";
}

/// <summary>
/// Writes the HTML Help Workshop project file (.hhp) including the [MAP] and [ALIAS]
/// sections that bind context IDs to topics, plus the .hhc/.hhk sitemaps and the C++
/// header consumed by the host application.
/// </summary>
public static class ChmProjectGenerator
{
    public static string GenerateHhp(HelpDocument document, ChmProjectNames names, IEnumerable<string> files)
    {
        var builder = new StringBuilder();

        builder.AppendLine("[OPTIONS]");
        builder.AppendLine("Compatibility=1.1 or later");
        builder.AppendLine("Compiled file=" + names.ChmFile);
        builder.AppendLine("Contents file=" + names.HhcFile);
        if (document.IndexEntries.Count > 0)
        {
            builder.AppendLine("Index file=" + names.HhkFile);
        }

        builder.AppendLine("Default topic=" + (document.DefaultTopicFileName ?? "index.html"));
        builder.AppendLine("Title=" + document.Title);
        builder.AppendLine("Language=" + MapLanguage(document.Language));
        builder.AppendLine("Display compile progress=No");
        builder.AppendLine("Full-text search=Yes");
        builder.AppendLine("Default Window=Main");
        builder.AppendLine();
        builder.AppendLine("[WINDOWS]");
        builder.AppendLine("Main=\"" + document.Title + "\",\"" + names.HhcFile + "\",\"" +
                           (document.IndexEntries.Count > 0 ? names.HhkFile : string.Empty) + "\",\"" +
                           (document.DefaultTopicFileName ?? "index.html") +
                           "\",,,,,,,,,0x23520,,0x384e,,,,,,,,0");
        builder.AppendLine();
        builder.AppendLine("[FILES]");
        foreach (var file in files)
        {
            builder.AppendLine(file.Replace('\\', '/'));
        }

        builder.AppendLine();
        builder.AppendLine("[ALIAS]");
        foreach (var page in document.Pages)
        {
            if (page.Symbol is not null)
            {
                builder.AppendLine($"{page.Symbol}={page.FileName}");
            }

            foreach (var anchor in page.Anchors)
            {
                // The compiler resolves an alias only to a topic file, so the symbol
                // targets the redirect stub rather than "page.html#anchor".
                builder.AppendLine($"{anchor.Symbol}={anchor.FileName}");
            }
        }

        builder.AppendLine();
        builder.AppendLine("[MAP]");
        foreach (var page in document.Pages)
        {
            if (page.Symbol is not null && page.ContextId.HasValue)
            {
                builder.AppendLine($"#define {page.Symbol} {page.ContextId.Value}");
            }

            foreach (var anchor in page.Anchors)
            {
                builder.AppendLine($"#define {anchor.Symbol} {anchor.ContextId}");
            }
        }

        builder.AppendLine();
        builder.AppendLine("[INFOTYPES]");
        return builder.ToString();
    }

    public static string GenerateHhc(HelpDocument document)
    {
        var builder = new StringBuilder();
        WriteSiteProperties(builder);
        builder.AppendLine("<UL>");

        foreach (var node in document.Toc)
        {
            AppendTocNode(builder, node, 1);
        }

        builder.AppendLine("</UL>");
        builder.AppendLine("</BODY>");
        builder.AppendLine("</HTML>");
        return builder.ToString();
    }

    private static void AppendTocNode(StringBuilder builder, TocNode node, int depth)
    {
        var indent = new string(' ', depth * 2);
        builder.Append(indent).AppendLine("<LI> <OBJECT type=\"text/sitemap\">");
        builder.Append(indent).AppendLine($"  <param name=\"Name\" value=\"{Escape(node.Title)}\">");
        if (!string.IsNullOrEmpty(node.Local))
        {
            builder.Append(indent).AppendLine($"  <param name=\"Local\" value=\"{Escape(node.Local)}\">");
        }

        builder.Append(indent).AppendLine("</OBJECT>");

        if (node.Children.Count > 0)
        {
            builder.Append(indent).AppendLine("<UL>");
            foreach (var child in node.Children)
            {
                AppendTocNode(builder, child, depth + 1);
            }

            builder.Append(indent).AppendLine("</UL>");
        }
    }

    public static string GenerateHhk(HelpDocument document)
    {
        var builder = new StringBuilder();
        WriteSiteProperties(builder);
        builder.AppendLine("<UL>");

        // HTML Help's index expects a keyword to appear once as a parent; repeated
        // keywords as sibling leaves make hh.exe fail with the "not enough memory"
        // error, so every keyword is emitted exactly once and duplicate targets are
        // collapsed.
        var groups = document.IndexEntries
            .GroupBy(e => e.Keyword, StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase);

        foreach (var group in groups)
        {
            var subKeywords = group
                .Where(e => !string.IsNullOrEmpty(e.SubKeyword))
                .GroupBy(e => e.SubKeyword!, StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var plainTargets = group
                .Where(e => string.IsNullOrEmpty(e.SubKeyword))
                .Select(e => (e.FileName, e.Anchor))
                .Distinct()
                .ToList();

            // The parent node is itself a link when the keyword has a plain target.
            var primary = plainTargets.FirstOrDefault();
            builder.AppendLine("  <LI> <OBJECT type=\"text/sitemap\">");
            builder.AppendLine($"    <param name=\"Name\" value=\"{Escape(group.Key)}\">");
            if (primary.FileName is not null)
            {
                builder.AppendLine($"    <param name=\"Local\" value=\"{Escape(Local(primary.FileName, primary.Anchor))}\">");
            }

            builder.AppendLine("  </OBJECT>");

            if (subKeywords.Count == 0)
            {
                // Extra plain targets become unnamed child links; with a single target
                // the parent link above already covers it.
                var extra = plainTargets.Skip(1).ToList();
                if (extra.Count > 0)
                {
                    builder.AppendLine("  <UL>");
                    foreach (var target in extra)
                    {
                        AppendIndexLeaf(builder, target.FileName, target.Anchor);
                    }

                    builder.AppendLine("  </UL>");
                }

                continue;
            }

            builder.AppendLine("  <UL>");
            foreach (var sub in subKeywords)
            {
                var target = sub.Select(e => (e.FileName, e.Anchor)).First();
                AppendIndexLeaf(builder, target.FileName, target.Anchor, sub.Key);
            }

            builder.AppendLine("  </UL>");
        }

        builder.AppendLine("</UL>");
        builder.AppendLine("</BODY>");
        builder.AppendLine("</HTML>");
        return builder.ToString();
    }

    private static void WriteSiteProperties(StringBuilder builder)
    {
        builder.AppendLine("<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML//EN\">");
        builder.AppendLine("<HTML>");
        builder.AppendLine("<BODY>");
        builder.AppendLine("<OBJECT type=\"text/site properties\">");
        builder.AppendLine("<param name=\"Window Styles\" value=\"0x23520\">");
        builder.AppendLine("</OBJECT>");
    }

    private static string Local(string fileName, string? anchor) =>
        string.IsNullOrEmpty(anchor) ? fileName : fileName + "#" + anchor;

    private static void AppendIndexLeaf(StringBuilder builder, string fileName, string? anchor, string? name = null)
    {
        builder.AppendLine("    <LI> <OBJECT type=\"text/sitemap\">");
        if (name is not null)
        {
            builder.AppendLine($"      <param name=\"Name\" value=\"{Escape(name)}\">");
        }

        builder.AppendLine($"      <param name=\"Local\" value=\"{Escape(Local(fileName, anchor))}\">");
        builder.AppendLine("    </OBJECT>");
    }

    /// <summary>
    /// Emits the header to include in a C++ project. Every symbol is a stable compile-time
    /// constant, so call sites never depend on generated numeric values.
    /// </summary>
    public static string GenerateHeader(HelpDocument document, ChmProjectNames names, string? chmRelativePath = null)
    {
        var guard = "_" + Sanitize(names.BaseName).ToUpperInvariant() + "_H_";
        var builder = new StringBuilder();

        builder.AppendLine("// Generato automaticamente da Word2Chm. Non modificare a mano.");
        builder.AppendLine("// Compila con: HtmlHelp(hwnd, L\"" + (chmRelativePath ?? names.ChmFile) + "\", HH_HELP_CONTEXT, IDH);");
        builder.AppendLine("#pragma once");
        builder.AppendLine();
        builder.AppendLine("#ifndef " + guard);
        builder.AppendLine("#define " + guard);
        builder.AppendLine();

        foreach (var page in document.Pages)
        {
            if (page.Symbol is not null && page.ContextId.HasValue)
            {
                builder.AppendLine($"#define {page.Symbol,-40} {page.ContextId.Value}   // {page.Title}");
            }

            foreach (var anchor in page.Anchors)
            {
                builder.AppendLine($"#define {anchor.Symbol,-40} {anchor.ContextId}   // {page.Title} > {anchor.Title}");
            }
        }

        builder.AppendLine();
        builder.AppendLine("#endif // " + guard);
        return builder.ToString();
    }

    private static string Sanitize(string value)
    {
        var chars = value.Select(c => char.IsLetterOrDigit(c) ? c : '_').ToArray();
        return new string(chars);
    }

    private static string MapLanguage(string language)
    {
        // HTML Help Workshop expects an LCID in hexadecimal, e.g. 0x410 for Italian.
        var name = language.Split('-', '_')[0].ToLowerInvariant();

        // Fast path for the common documentation languages, which also works under
        // invariant globalization where CultureInfo lookups are unavailable.
        if (KnownLcids.TryGetValue(name, out var known))
        {
            return known;
        }

        var culture = CultureInfo.GetCultures(CultureTypes.AllCultures)
            .FirstOrDefault(c => c.TwoLetterISOLanguageName == name && c.LCID > 0 && c.LCID != 0x1000);

        if (culture is null)
        {
            return "0x409 English (United States)";
        }

        return $"0x{culture.LCID:X} {culture.EnglishName}";
    }

    private static readonly Dictionary<string, string> KnownLcids = new(StringComparer.OrdinalIgnoreCase)
    {
        ["en"] = "0x409 English (United States)",
        ["it"] = "0x410 Italian (Italy)",
        ["de"] = "0x407 German (Germany)",
        ["fr"] = "0x40C French (France)",
        ["es"] = "0x40A Spanish (Spain)",
        ["pt"] = "0x816 Portuguese (Portugal)",
        ["nl"] = "0x413 Dutch (Netherlands)",
        ["pl"] = "0x415 Polish (Poland)",
        ["ru"] = "0x419 Russian (Russia)",
        ["zh"] = "0x804 Chinese (Simplified, PRC)",
        ["ja"] = "0x411 Japanese (Japan)",
        ["ko"] = "0x412 Korean (Korea)",
    };

    private static string Escape(string value)
    {
        // hhc.exe reads .hhc/.hhk as ANSI and mishandles raw UTF-8, which can leave a
        // CHM that compiles but refuses to open. Emitting non-ASCII characters as
        // numeric HTML entities keeps these files pure ASCII and still renders the
        // accented text correctly in the viewer.
        var builder = new StringBuilder(value.Length);

        foreach (var ch in value)
        {
            switch (ch)
            {
                case '&':
                    builder.Append("&amp;");
                    break;
                case '"':
                    builder.Append("&quot;");
                    break;
                case '<':
                    builder.Append("&lt;");
                    break;
                case '>':
                    builder.Append("&gt;");
                    break;
                default:
                    if (ch > 127)
                    {
                        builder.Append("&#").Append((int)ch).Append(';');
                    }
                    else
                    {
                        builder.Append(ch);
                    }

                    break;
            }
        }

        return builder.ToString();
    }
}
