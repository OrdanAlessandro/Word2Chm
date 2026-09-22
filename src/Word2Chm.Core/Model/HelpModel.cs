namespace Word2Chm.Core.Model;

/// <summary>Where a Word bookmark landed after page splitting.</summary>
public sealed record BookmarkTarget(string PageFileName, string Anchor);

/// <summary>
/// The complete intermediate representation of a compiled help project, produced
/// by the DOCX parser and consumed by the generators.
/// </summary>
public sealed class HelpDocument
{
    public string Title { get; set; } = "Guida";

    /// <summary>Culture code (e.g. "it-IT") written to the CHM language field.</summary>
    public string Language { get; set; } = "it-IT";

    /// <summary>Base numeric value assigned to the first context ID.</summary>
    public int DefaultContextId { get; set; } = 1000;

    public List<HelpPage> Pages { get; } = new();

    public List<TocNode> Toc { get; } = new();

    public List<IndexEntry> IndexEntries { get; } = new();

    /// <summary>Resolved Word bookmark name to page/anchor, used for internal links.</summary>
    public Dictionary<string, BookmarkTarget> BookmarkTargets { get; } = new(StringComparer.Ordinal);

    /// <summary>File name of the page shown when the CHM opens without a context ID.</summary>
    public string? DefaultTopicFileName => Pages.Count > 0 ? Pages[0].FileName : null;

    /// <summary>
    /// Non-fatal problems found while building the project, such as a symbol declared
    /// twice. Surfacing these avoids the silent disappearance of identifiers the user
    /// expects to find in the header. Context IDs on sub-headings are no longer reported
    /// here: they are bound to an in-page anchor through <see cref="HelpPage.Anchors"/>.
    /// </summary>
    public List<string> Warnings { get; } = new();
}

/// <summary>A single HTML output file, anchored to a Help 1 context ID.</summary>
public sealed class HelpPage
{
    /// <summary>Symbolic identifier declared in the document, e.g. <c>IDH_INSTALLAZIONE</c>.</summary>
    public string? Symbol { get; set; }

    /// <summary>Numeric context ID written to the [MAP] section. Null when no symbol was declared.</summary>
    public int? ContextId { get; set; }

    /// <summary>Visible heading text, with the context ID marker stripped.</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>Output file name relative to the project root.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>Heading level that opened this page (1 for "Heading 1").</summary>
    public int Level { get; set; } = 1;

    /// <summary>Anchor generated for this page's own heading.</summary>
    public string? Anchor { get; set; }

    public List<DocumentBlock> Blocks { get; } = new();

    /// <summary>Index entries that point at this page.</summary>
    public List<IndexEntry> IndexEntries { get; } = new();

    /// <summary>
    /// Additional context IDs on sub-headings of this page. A Help 1 ID normally maps to a
    /// topic file, but <c>[ALIAS]</c> also accepts <c>file.htm#anchor</c>, so a marker on a
    /// heading below the page level resolves to an in-page anchor instead of being dropped.
    /// </summary>
    public List<HelpAnchor> Anchors { get; } = new();
}

/// <summary>A symbolic context ID bound to an anchor inside a page.</summary>
public sealed class HelpAnchor
{
    /// <summary>Symbolic identifier declared in the document, e.g. <c>IDH_AXES</c>.</summary>
    public string Symbol { get; set; } = string.Empty;

    /// <summary>Numeric context ID written to the [MAP] section.</summary>
    public int ContextId { get; set; }

    /// <summary>Anchor generated for the heading that declared the symbol.</summary>
    public string Anchor { get; set; } = string.Empty;

    /// <summary>Visible heading text, used for the generated header comment.</summary>
    public string Title { get; set; } = string.Empty;
}

/// <summary>An entry of the table of contents tree, mirroring document heading levels.</summary>
public sealed class TocNode
{
    public string Title { get; set; } = string.Empty;
    public string? Local { get; set; }
    public int Level { get; set; }
    public List<TocNode> Children { get; } = new();
}

public sealed class IndexEntry
{
    public string Keyword { get; set; } = string.Empty;
    public string? SubKeyword { get; set; }

    /// <summary>Page file name the entry links to.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>Optional anchor within the page.</summary>
    public string? Anchor { get; set; }
}
