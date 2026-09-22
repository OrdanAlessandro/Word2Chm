using Word2Chm.Core.Common;
using Word2Chm.Core.Docx;
using Word2Chm.Core.Model;

namespace Word2Chm.Core.Generation;

public sealed class BuildOptions
{
    /// <summary>Base numeric value for auto-assigned context IDs.</summary>
    public int DefaultContextId { get; set; } = 1000;

    /// <summary>Heading level that starts a new HTML page.</summary>
    public int PageLevel { get; set; } = 1;

    /// <summary>Deepest heading level included in the table of contents.</summary>
    public int MaxTocLevel { get; set; } = 6;

    public string Language { get; set; } = "it-IT";

    /// <summary>Optional external overrides, keyed by symbol name.</summary>
    public Dictionary<string, int> ContextIdOverrides { get; set; } = new(StringComparer.Ordinal);
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
            DefaultContextId = options.DefaultContextId,
        };

        var usedIds = new HashSet<int>();
        var tocBuilder = new TocBuilder(options.MaxTocLevel);

        // Index entries reference the nearest preceding heading; we record it while scanning.
        var headingTargets = new Dictionary<DocumentBlock, (HelpPage Page, string Anchor)>();
        var pendingIndexEntries = new List<(IndexKeyword Keyword, DocumentBlock Block)>();
        var nextId = options.DefaultContextId;

        HelpPage? current = null;

        foreach (var block in parsed.Blocks)
        {
            if (block is HeadingBlock heading)
            {
                if (heading.Level <= options.PageLevel || current is null)
                {
                    current = CreatePage(heading, document.Pages.Count, options, usedIds, ref nextId);
                    document.Pages.Add(current);
                }
                else if (!string.IsNullOrEmpty(heading.Symbol))
                {
                    // A Help 1 context ID resolves to a topic file, and the compiler does not
                    // accept an anchor in that file name, so a marker on a sub-heading is
                    // bound to a small redirect topic that forwards to the page anchor.
                    heading.Anchor = slugger.Slug(heading.Title);
                    var id = ResolveContextId(heading.Symbol, heading.ExplicitId, options, usedIds, ref nextId);
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
                    current.Blocks.Add(block);
                    continue;
                }

                heading.Anchor = slugger.Slug(heading.Title);
                var target = (current!, heading.Anchor);
                headingTargets[block] = target;
                RegisterBookmarks(block, target);
                tocBuilder.Add(heading.Level, heading.Title, current.FileName + "#" + heading.Anchor);
                RegisterIndexKeywords(block, heading.IndexKeywords, pendingIndexEntries);
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
        BuildOptions options,
        HashSet<int> usedIds,
        ref int nextId)
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
            page.ContextId = ResolveContextId(page.Symbol, heading.ExplicitId, options, usedIds, ref nextId);
        }

        return page;
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

    private static int ResolveContextId(
        string symbol,
        int? explicitId,
        BuildOptions options,
        HashSet<int> usedIds,
        ref int nextId)
    {
        var id = explicitId
                 ?? (options.ContextIdOverrides.TryGetValue(symbol, out var over) ? over : (int?)null)
                 ?? TakeNext(ref nextId, usedIds);

        if (!usedIds.Add(id))
        {
            throw new InvalidOperationException(
                $"L'ID di contesto {id} per il simbolo '{symbol}' è già usato da un altro simbolo.");
        }

        return id;
    }

    private static int TakeNext(ref int nextId, HashSet<int> usedIds)
    {
        while (usedIds.Contains(nextId))
        {
            nextId++;
        }

        var value = nextId;
        nextId = value + 1;
        return value;
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
