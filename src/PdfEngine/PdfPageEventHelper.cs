namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// PDF页面事件辅助类
/// </summary>
public class PdfPageEventHelper
{
    public virtual void OnOpenDocument(PdfWriter writer, PdfDocument document) { }
    public virtual void OnStartPage(PdfWriter writer, PdfDocument document) { }
    public virtual void OnEndPage(PdfWriter writer, PdfDocument document) { }
    public virtual void OnCloseDocument(PdfWriter writer, PdfDocument document) { }
}

/// <summary>
/// 节追踪器
/// </summary>
public class SectionTracker : PdfPageEventHelper
{
    public int CurrentSection { get; set; }
    public List<int> PageSections { get; } = [];

    public override void OnStartPage(PdfWriter writer, PdfDocument document)
    {
        PageSections.Add(CurrentSection);
    }
}

/// <summary>
/// 书签追踪器
/// </summary>
public class BookmarkTracker : PdfPageEventHelper
{
    private readonly PdfWriter _writer;
    private PdfOutline? _rootOutline;
    private readonly Dictionary<int, PdfOutline> _outlineByLevel = [];

    public BookmarkTracker(PdfWriter writer)
    {
        _writer = writer;
    }

    public override void OnOpenDocument(PdfWriter writer, PdfDocument document)
    {
        _rootOutline = new PdfOutline(null, new PdfDestination(PdfDestination.XYZ), "Bookmarks", true);
        writer.SetRootOutline(_rootOutline);
    }

    public void AddHeadingBookmark(string title, int level)
    {
        if (_rootOutline == null) return;

        level = Math.Clamp(level, 1, 6);

        var dest = new PdfDestination(PdfDestination.XYZ, -1, 0, 0);
        var parent = level == 1 ? _rootOutline : _outlineByLevel.GetValueOrDefault(level - 1, _rootOutline);
        var outline = new PdfOutline(parent, dest, title, level <= 2);

        _outlineByLevel[level] = outline;
    }

    public void AddBookmark(string name)
    {
        if (_rootOutline == null) return;

        var dest = new PdfDestination(PdfDestination.XYZ, -1, 0, 0);
        new PdfOutline(_rootOutline, dest, name, false);
    }
}

/// <summary>
/// 组合页面事件
/// </summary>
public class CombinedPageEvent : PdfPageEventHelper
{
    private readonly PdfPageEventHelper[] _events;

    public CombinedPageEvent(params PdfPageEventHelper[] events)
    {
        _events = events;
    }

    public override void OnOpenDocument(PdfWriter writer, PdfDocument document)
    {
        foreach (var e in _events) e.OnOpenDocument(writer, document);
    }

    public override void OnStartPage(PdfWriter writer, PdfDocument document)
    {
        foreach (var e in _events) e.OnStartPage(writer, document);
    }

    public override void OnEndPage(PdfWriter writer, PdfDocument document)
    {
        foreach (var e in _events) e.OnEndPage(writer, document);
    }

    public override void OnCloseDocument(PdfWriter writer, PdfDocument document)
    {
        foreach (var e in _events) e.OnCloseDocument(writer, document);
    }
}
