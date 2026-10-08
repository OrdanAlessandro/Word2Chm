using Word2Chm.Core.Common;
using Word2Chm.Core.Docx;
using Word2Chm.Core.Model;

namespace Word2Chm.Core.Generation;

public sealed class BuildOptions
{
    /// <summary>
    /// Deepest heading level that opens a new page, shared by the core and the GUI so the
    /// two can never disagree about the default.
    /// </summary>
    public const int DefaultPageLevel = 6;

    /// <summary>
    /// Body text size in points, shared by the core and the GUI so the two can never
    /// disagree about the default. Headings keep their own relative sizes and are not
    /// affected, which is what lets the body be enlarged on its own.
    /// </summary>
    public const double DefaultBodyFontSizePt = 10.5;

    /// <summary>
    /// Deepest heading level that opens a new HTML page. It defaults to
    /// <see cref="DefaultPageLevel"/> so every heading that appears in the table of
    /// contents gets its own topic; the viewer would otherwise show a whole chapter as
    /// one long page and the menu would be the only way to move inside it.
    /// </summary>
    public int PageLevel { get; set; } = DefaultPageLevel;

    /// <summary>
    /// Size of the normal text, in points. Headings are sized in em and keep their
    /// proportions, so only paragraphs, lists and tables follow this value.
    /// </summary>
    public double BodyFontSizePt { get; set; } = DefaultBodyFontSizePt;

    /// <summary>Deepest heading level included in the table of contents.</summary>
    public int MaxTocLevel { get; set; } = 6;

    public string Language { get; set; } = "it-IT";

    /// <summary>Footer text written to every page; see <see cref="HelpDocument.DefaultFooter"/>.</summary>
    public string Footer { get; set; } = HelpDocument.DefaultFooter;

    /// <summary>
    /// Header supplying the numeric context IDs, keyed by symbol name. The document only
    /// names a symbol; the number always comes from here, so the IDs stay identical across
    /// language editions that share one header.
    /// </summary>
    public ContextIdHeader? ContextIds { get; set; }

    /// <summary>
    /// Symbols referenced by the document but absent from <see cref="ContextIds"/>. Filled
    /// while building; the conversion still completes so the user can see the whole list at
    /// once instead of fixing one symbol per run.
    /// </summary>
    public List<string> MissingSymbols { get; } = new();
}

/// <summary>
/// Turns the flat parsed document into a paged <see cref="HelpDocument"/>: headings at
/// the configured level start a new HTML page, anchors are generated for headings and
/// bookmarks, the table of contents and the index are assembled, and Word bookmarks are
/// resolved to page/anchor pairs.
/// </summary>
public sealed class HelpProjectBuilder
{
    private readonly Dictionary<string, BookmarkTarget> _bookmarks = new(StringComparer.Ordinal);

    public HelpDocument Build(ParsedDocument parsed, BuildOptions options)
    {
        var slugger = new Slugger();
        var document = new HelpDocument
        {
            Title = parsed.Title,
            Language = options.Language,
            Footer = options.Footer,
        };

        var tocBuilder = new TocBuilder(options.MaxTocLevel);

        // Index entries reference the nearest preceding heading; we record it while scanning.
        var headingTargets = new Dictionary<DocumentBlock, (HelpPage Page, string Anchor)>();
        var pendingIndexEntries = new List<(IndexKeyword Keyword, DocumentBlock Block)>();

        HelpPage? current = null;

        // Chain of headings seen since the last page break, used to build breadcrumbs.
        var headingStack = new List<(int Level, string Title, string? Local)>();

        foreach (var block in parsed.Blocks)
        {
            if (block is HeadingBlock heading)
            {
                while (headingStack.Count > 0 && headingStack[^1].Level >= heading.Level)
                {
                    headingStack.RemoveAt(headingStack.Count - 1);
                }

                var isPageStart = heading.Level <= options.PageLevel || current is null;
                if (isPageStart)
                {
                    current = CreatePage(heading, document.Pages.Count, options);
                    SetAncestors(current, headingStack.Select(h => (h.Title, h.Local)).ToList());
                    document.Pages.Add(current);

                    heading.Anchor = slugger.Slug(heading.Title);
                    var pageTarget = (current, heading.Anchor);
                    headingTargets[block] = pageTarget;
                    RegisterBookmarks(block, pageTarget);
                    tocBuilder.Add(heading.Level, heading.Title, current.FileName + "#" + heading.Anchor);
                    RegisterIndexKeywords(block, heading.IndexKeywords, pendingIndexEntries);
                    headingStack.Add((heading.Level, heading.Title, current.FileName + "#" + heading.Anchor));
                    current.Blocks.Add(block);
                    continue;
                }

                if (!string.IsNullOrEmpty(heading.Symbol))
                {
                    // A Help 1 context ID resolves to a topic file, and the compiler does not
                    // accept an anchor in that file name, so a marker on a sub-heading is
                    // bound to a small redirect topic that forwards to the page anchor.
                    heading.Anchor = slugger.Slug(heading.Title);
                    var id = ResolveContextId(heading.Symbol, options);
                    var anchorTarget = $"{current.FileName}#{heading.Anchor}";
                    current.Anchors.Add(new HelpAnchor
                    {
                        Symbol = heading.Symbol!,
                        ContextId = id,
                        Anchor = heading.Anchor,
                        Title = heading.Title,
                        FileName = AnchorFileName(current.FileName, heading.Anchor),
                        Target = anchorTarget,
                    });

                    var aliasTarget = (current, heading.Anchor);
                    headingTargets[block] = aliasTarget;
                    RegisterBookmarks(block, aliasTarget);
                    tocBuilder.Add(heading.Level, heading.Title, current.FileName + "#" + heading.Anchor);
                    RegisterIndexKeywords(block, heading.IndexKeywords, pendingIndexEntries);
                    headingStack.Add((heading.Level, heading.Title, anchorTarget));
                    current.Blocks.Add(block);
                    continue;
                }

                heading.Anchor = slugger.Slug(heading.Title);
                var target = (current, heading.Anchor);
                headingTargets[block] = target;
                RegisterBookmarks(block, target);
                tocBuilder.Add(heading.Level, heading.Title, current.FileName + "#" + heading.Anchor);
                RegisterIndexKeywords(block, heading.IndexKeywords, pendingIndexEntries);
                headingStack.Add((heading.Level, heading.Title, current.FileName + "#" + heading.Anchor));
                current.Blocks.Add(block);
                continue;
            }

            current ??= CreateImplicitPage(document);
            if (block.Bookmarks.Count > 0)
            {
                block.Anchor ??= slugger.Slug("bookmark");
                RegisterBookmarks(block, (current, block.Anchor));
            }

            if (block is ParagraphBlock { IndexKeywords.Count: > 0 } paragraph)
            {
                foreach (var keyword in paragraph.IndexKeywords)
                {
                    pendingIndexEntries.Add((keyword, block));
                }
            }

            current.Blocks.Add(block);
        }

        tocBuilder.Apply(document);

        // Headings that have children are section labels the menu expands, so their entry
        // must keep pointing at a page of its own; demoting one would send its click into
        // whichever page happens to precede it.
        var localsWithChildren = new HashSet<string>(StringComparer.Ordinal);
        CollectLocalsWithChildren(document.Toc, localsWithChildren);

        // Pages that hold nothing but their own heading are structural labels ("Assi",
        // "Utilita"): Word numbers every heading level 1-3 as a chapter, so without this
        // the menu fills with one-line pages. A leaf label is demoted to the preceding
        // page, which keeps the heading as an anchor there, while an explicit context ID
        // or index keyword still forces a real page because something links to it by name.
        CollapseContentlessPages(document, headingTargets, localsWithChildren);

        foreach (var (keyword, block) in pendingIndexEntries)
        {
            var target = FindNearestHeading(parsed, block, headingTargets);
            if (target is null)
            {
                continue;
            }

            var (page, anchor) = target.Value;
            var entry = new IndexEntry
            {
                Keyword = keyword.Keyword,
                SubKeyword = keyword.SubKeyword,
                FileName = page.FileName,
                Anchor = anchor == page.Anchor ? null : anchor,
            };

            document.IndexEntries.Add(entry);
            page.IndexEntries.Add(entry);
        }

        parsed.BookmarkTargets.Clear();
        foreach (var (name, target) in _bookmarks)
        {
            parsed.BookmarkTargets[name] = target;
            document.BookmarkTargets[name] = target;
        }

        return document;
    }

    /// <summary>
    /// Demotes pages whose only block is their own heading into the nearest preceding page
    /// that has content. The heading stays in the output as an anchor, so links and the
    /// table of contents keep working, but the menu no longer lists one-line pages.
    /// Pages that must stay reachable on their own are left untouched: the ones named by a
    /// context ID, an index keyword or a bookmark, and the section labels that
    /// <paramref name="localsWithChildren"/> lists, whose menu entry would otherwise open
    /// an unrelated page.
    /// </summary>
    private static void CollapseContentlessPages(
        HelpDocument document,
        Dictionary<DocumentBlock, (HelpPage Page, string Anchor)> headingTargets,
        HashSet<string> localsWithChildren)
    {
        var survivors = new HashSet<HelpPage>();
        for (var i = 0; i < document.Pages.Count; i++)
        {
            var page = document.Pages[i];
            var onlyHeading = page.Blocks.Count == 1 && page.Blocks[0] is HeadingBlock;
            var hasKeyword = page.Blocks.Count == 1 &&
                             page.Blocks[0] is HeadingBlock { IndexKeywords.Count: > 0 };
            var isSectionLabel = page.Blocks is [HeadingBlock heading] &&
                                 heading.Anchor is not null &&
                                 localsWithChildren.Contains(page.FileName + "#" + heading.Anchor);
            if (i == 0 || !onlyHeading || hasKeyword || page.Symbol is not null || isSectionLabel)
            {
                survivors.Add(page);
            }
        }

        var moved = new Dictionary<string, string>(StringComparer.Ordinal);
        var collapsed = new List<(HelpPage Page, HelpPage Parent, DocumentBlock Block)>();
        for (var i = 0; i < document.Pages.Count; i++)
        {
            var page = document.Pages[i];
            if (survivors.Contains(page))
            {
                continue;
            }

            HelpPage? parent = null;
            for (var j = i - 1; j >= 0; j--)
            {
                if (survivors.Contains(document.Pages[j]))
                {
                    parent = document.Pages[j];
                    break;
                }
            }

            if (parent is null)
            {
                continue;
            }

            var block = page.Blocks[0];
            parent.Blocks.Add(block);
            collapsed.Add((page, parent, block));
            moved[page.FileName] = parent.FileName;
        }

        foreach (var (page, parent, block) in collapsed)
        {
            document.Pages.Remove(page);

            if (headingTargets.Remove(block, out var target))
            {
                headingTargets[block] = (parent, target.Anchor);
            }

            // Anything that resolved to the removed page must follow the heading.
            if (block is HeadingBlock heading && heading.Anchor is not null)
            {
                foreach (var name in block.Bookmarks)
                {
                    document.BookmarkTargets[name] = new BookmarkTarget(parent.FileName, heading.Anchor);
                }
            }
        }

        if (moved.Count == 0)
        {
            return;
        }

        // Table of contents entries and breadcrumbs still name the removed files.
        foreach (var node in document.Toc)
        {
            RemapTocNode(node, moved);
        }

        foreach (var page in document.Pages)
        {
            for (var i = 0; i < page.Ancestors.Count; i++)
            {
                var (title, local) = page.Ancestors[i];
                page.Ancestors[i] = (title, RemapLocal(local, moved));
            }
        }
    }

    /// <summary>
    /// Records the "file.htm#anchor" of every table-of-contents entry that has children,
    /// i.e. every section label the menu expands.
    /// </summary>
    private static void CollectLocalsWithChildren(IEnumerable<TocNode> nodes, HashSet<string> into)
    {
        foreach (var node in nodes)
        {
            if (node.Children.Count > 0 && !string.IsNullOrEmpty(node.Local))
            {
                into.Add(node.Local);
            }

            CollectLocalsWithChildren(node.Children, into);
        }
    }

    private static void RemapTocNode(TocNode node, Dictionary<string, string> moved)
    {
        node.Local = RemapLocal(node.Local, moved);
        foreach (var child in node.Children)
        {
            RemapTocNode(child, moved);
        }
    }

    private static string? RemapLocal(string? local, Dictionary<string, string> moved)
    {
        if (string.IsNullOrEmpty(local))
        {
            return local;
        }

        var separator = local.IndexOf('#');
        var file = separator < 0 ? local : local[..separator];
        var anchor = separator < 0 ? string.Empty : local[separator..];
        return moved.TryGetValue(file, out var replacement) ? replacement + anchor : local;
    }

    private void RegisterBookmarks(DocumentBlock block, (HelpPage Page, string Anchor) target)
    {
        foreach (var name in block.Bookmarks)
        {
            _bookmarks[name] = new BookmarkTarget(target.Page.FileName, target.Anchor);
        }
    }

    /// <summary>
    /// Builds the file name of the redirect topic for an anchored context ID. The page
    /// slug is reused with an <c>-id</c> suffix so the stub stays recognisable next to
    /// the page it forwards to.
    /// </summary>
    private static string AnchorFileName(string pageFileName, string anchor)
    {
        var stem = Path.GetFileNameWithoutExtension(pageFileName);
        return stem + "-" + anchor + "-id.html";
    }

    private static void RegisterIndexKeywords(
        DocumentBlock block,
        IEnumerable<IndexKeyword> keywords,
        List<(IndexKeyword Keyword, DocumentBlock Block)> pending)
    {
        foreach (var keyword in keywords)
        {
            pending.Add((keyword, block));
        }
    }

    private static HelpPage CreatePage(
        HeadingBlock heading,
        int index,
        BuildOptions options)
    {
        var page = new HelpPage
        {
            Title = heading.Title,
            Level = heading.Level,
            FileName = GenerateFileName(index, heading.Title),
            Symbol = string.IsNullOrEmpty(heading.Symbol) ? null : heading.Symbol,
        };

        if (page.Symbol is not null)
        {
            page.ContextId = ResolveContextId(page.Symbol, options);
        }

        return page;
    }

    /// <summary>
    /// Records the heading chain that leads to a page, so the skin can render a breadcrumb.
    /// Ancestors are kept as a stack of the headings seen since the last page break.
    /// </summary>
    private static void SetAncestors(
        HelpPage page,
        IReadOnlyList<(string Title, string? Local)> stack)
    {
        page.Ancestors.AddRange(stack);
    }

    private static HelpPage CreateImplicitPage(HelpDocument document)
    {
        var title = document.Pages.Count == 0 ? document.Title : "Contenuto";
        var page = new HelpPage
        {
            Title = title,
            Level = 1,
            FileName = GenerateFileName(document.Pages.Count, title),
        };

        document.Pages.Add(page);
        return page;
    }

    /// <summary>
    /// Looks the symbol up in the supplied header. A symbol the header does not define is
    /// recorded and reported instead of stopping the run: the conversion still produces a
    /// complete CHM, so the user fixes every missing ID in one pass. The topic keeps its
    /// file and alias but carries no [MAP] entry, which the compiler accepts.
    /// </summary>
    private static int? ResolveContextId(string symbol, BuildOptions options)
    {
        if (options.ContextIds is not null && options.ContextIds.TryGetId(symbol, out var id))
        {
            return id;
        }

        if (!options.MissingSymbols.Contains(symbol, StringComparer.Ordinal))
        {
            options.MissingSymbols.Add(symbol);
        }

        return null;
    }

    private static (HelpPage Page, string Anchor)? FindNearestHeading(
        ParsedDocument parsed,
        DocumentBlock block,
        Dictionary<DocumentBlock, (HelpPage Page, string Anchor)> headingTargets)
    {
        (HelpPage Page, string Anchor)? last = null;

        foreach (var candidate in parsed.Blocks)
        {
            if (candidate == block)
            {
                return last;
            }

            if (headingTargets.TryGetValue(candidate, out var target))
            {
                last = target;
            }
        }

        return last;
    }

    private static string GenerateFileName(int index, string title)
    {
        var slug = new string(Slugger.Deaccent(title).Where(char.IsLetterOrDigit).ToArray())
            .ToLowerInvariant();

        if (slug.Length == 0)
        {
            slug = "pagina";
        }

        return $"{index:D3}-{slug}.html";
    }
}
