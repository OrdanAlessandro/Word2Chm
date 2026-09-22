using Word2Chm.Core.Model;

namespace Word2Chm.Core.Generation;

/// <summary>
/// Assembles the table of contents tree from a flat sequence of heading levels,
/// keeping relative depth so that skipped levels do not nest incorrectly.
/// </summary>
internal sealed class TocBuilder
{
    private readonly int _maxLevel;
    private readonly Stack<TocNode> _stack = new();
    private readonly List<(TocNode Node, int Level)> _roots = new();

    public TocBuilder(int maxLevel) => _maxLevel = maxLevel;

    public void Add(int level, string title, string local)
    {
        if (level > _maxLevel)
        {
            return;
        }

        var node = new TocNode { Title = title, Level = level, Local = local };

        while (_stack.Count > 0 && _stack.Peek().Level >= level)
        {
            _stack.Pop();
        }

        if (_stack.Count > 0)
        {
            _stack.Peek().Children.Add(node);
        }
        else
        {
            _roots.Add((node, level));
        }

        _stack.Push(node);
    }

    public void Apply(HelpDocument document)
    {
        foreach (var (node, _) in _roots)
        {
            document.Toc.Add(node);
        }
    }
}
