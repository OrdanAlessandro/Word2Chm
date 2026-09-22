using Word2Chm.Core.Model;
using BookmarkTarget = Word2Chm.Core.Model.BookmarkTarget;

namespace Word2Chm.Core.Docx;

/// <summary>Binary media extracted from the DOCX package.</summary>
public sealed class ImageResource
{
    public required string FileName { get; init; }
    public required byte[] Content { get; init; }
    public string ContentType { get; init; } = "image/png";
}

/// <summary>
/// Raw result of reading a DOCX file: a flat block sequence plus the bookkeeping
/// needed later to resolve bookmarks, images and index entries.
/// </summary>
public sealed class ParsedDocument
{
    public string Title { get; set; } = string.Empty;

    public List<DocumentBlock> Blocks { get; } = new();

    public List<ImageResource> Images { get; } = new();

    /// <summary>Maps a Word bookmark name to the generated anchor that represents it.</summary>
    public Dictionary<string, string> BookmarkAnchors { get; } = new(StringComparer.Ordinal);

    /// <summary>Maps a Word bookmark name to the page and anchor it belongs to.</summary>
    public Dictionary<string, BookmarkTarget> BookmarkTargets { get; } = new(StringComparer.Ordinal);
}
