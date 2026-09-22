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
        }

        builder.AppendLine();
        builder.AppendLine("[MAP]");
        foreach (var page in document.Pages)
        {
            if (page.Symbol is not null && page.ContextId.HasValue)
            {
                builder.AppendLine($"#define {page.Symbol} {page.ContextId.Value}");
            }
        }

        builder.AppendLine();
        builder.AppendLine("[INFOTYPES]");
        return builder.ToString();
    }

    public static string GenerateHhc(HelpDocument document)
    {
        var builder = new StringBuilder();
        builder.AppendLine("<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML//EN\">");
        builder.AppendLine("<HTML>");
        builder.AppendLine("<BODY>");
        builder.AppendLine("<OBJECT type=\"text/site properties\">");
        builder.AppendLine("<param name=\"Window Styles\" value=\"0x23520\">");
        builder.AppendLine("</OBJECT>");
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
        builder.AppendLine("<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML//EN\">");
        builder.AppendLine("<HTML>");
        builder.AppendLine("<BODY>");
        builder.AppendLine("<OBJECT type=\"text/site properties\">");
        builder.AppendLine("<param name=\"Window Styles\" value=\"0x23520\">");
        builder.AppendLine("</OBJECT>");
        builder.AppendLine("<UL>");

        // Group entries by primary keyword, then nest the sub-keywords.
        var groups = document.IndexEntries
            .GroupBy(e => e.Keyword, StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase);

        foreach (var group in groups)
        {
            var entries = group.ToList();
            var plain = entries.Where(e => string.IsNullOrEmpty(e.SubKeyword)).ToList();
            var withSub = entries.Where(e => !string.IsNullOrEmpty(e.SubKeyword)).ToList();

            if (withSub.Count == 0)
            {
                // A keyword with several plain targets becomes a parent with one leaf each.
                if (entries.Count > 1)
                {
                    builder.AppendLine("  <LI> <OBJECT type=\"text/sitemap\">");
                    builder.AppendLine($"    <param name=\"Name\" value=\"{Escape(group.Key)}\">");
                    builder.AppendLine("  </OBJECT>");
                    builder.AppendLine("  <UL>");
                    foreach (var entry in entries)
                    {
                        AppendIndexLeaf(builder, entry.FileName, entry.Anchor);
                    }

                    builder.AppendLine("  </UL>");
                }
                else
                {
                    AppendIndexLeaf(builder, entries[0].FileName, entries[0].Anchor, group.Key);
                }

                continue;
            }

            builder.AppendLine("  <LI> <OBJECT type=\"text/sitemap\">");
            builder.AppendLine($"    <param name=\"Name\" value=\"{Escape(group.Key)}\">");
            builder.AppendLine("  </OBJECT>");
            builder.AppendLine("  <UL>");

            // Plain targets sit alongside the sub-keywords so no entry is dropped.
            foreach (var entry in plain)
            {
                AppendIndexLeaf(builder, entry.FileName, entry.Anchor, group.Key);
            }

            foreach (var entry in withSub)
            {
                AppendIndexLeaf(builder, entry.FileName, entry.Anchor, entry.SubKeyword!);
            }

            builder.AppendLine("  </UL>");
        }

        builder.AppendLine("</UL>");
        builder.AppendLine("</BODY>");
        builder.AppendLine("</HTML>");
        return builder.ToString();
    }

    private static void AppendIndexLeaf(StringBuilder builder, string fileName, string? anchor, string? name = null)
    {
        var local = fileName + (string.IsNullOrEmpty(anchor) ? string.Empty : "#" + anchor);
        builder.AppendLine("    <LI> <OBJECT type=\"text/sitemap\">");
        if (name is not null)
        {
            builder.AppendLine($"      <param name=\"Name\" value=\"{Escape(name)}\">");
        }

        builder.AppendLine($"      <param name=\"Local\" value=\"{Escape(local)}\">");
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

    private static string Escape(string value) => value
        .Replace("&", "&amp;")
        .Replace("\"", "&quot;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;");
}
