namespace Word2Chm.Core.Common;

/// <summary>
/// Raised when the source .docx cannot be read: the file is missing, locked by Word, or
/// not a Word document at all. Wrapping those cases in one type lets the UI show a single
/// clear message instead of leaking framework exceptions such as FileFormatException.
/// </summary>
public sealed class DocxUnreadableException : Exception
{
    public DocxUnreadableException(string message, Exception? inner = null)
        : base(message, inner)
    {
    }
}
