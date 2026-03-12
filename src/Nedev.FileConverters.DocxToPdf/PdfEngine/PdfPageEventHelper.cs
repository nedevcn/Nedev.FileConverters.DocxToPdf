namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

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
    /// <summary>
    /// Page event that delegates to a HeaderFooterRenderer instance.
    /// </summary>
    public class HeaderFooterPageEvent : PdfPageEventHelper
    {
        private readonly HeaderFooterRenderer _renderer;
        private readonly SectionTracker _tracker;
        private readonly Dictionary<int, SectionPageSettings> _settings;

        public HeaderFooterPageEvent(HeaderFooterRenderer renderer,
                                      SectionTracker tracker,
                                      Dictionary<int, SectionPageSettings> settings)
        {
            _renderer = renderer;
            _tracker = tracker;
            _settings = settings;
        }

        public override void OnEndPage(PdfWriter writer, PdfDocument document)
        {
            int p = document.PageNumber;
            int total = document.PageNumber;
            int sectionIndex = _tracker.PageSections.Count >= p ? _tracker.PageSections[p - 1] : 0;
            int pageInSection = _tracker.PageSections.Take(p).Count(i => i == sectionIndex);
            int totalInSection = _tracker.PageSections.Count(i => i == sectionIndex);
            var pageSize = document.PageSize;
            var settings = _settings.GetValueOrDefault(sectionIndex) ?? new SectionPageSettings();
            var cb = writer.DirectContent;
            _renderer.Render(cb, pageSize, p, total, sectionIndex, pageInSection, totalInSection, settings);
        }
    }

    /// <summary>
    /// Page event that renders a simple text watermark centered on each page.
    /// </summary>
    public class WatermarkPageEvent : PdfPageEventHelper
    {
        private readonly WatermarkOptions _options;

        public WatermarkPageEvent(WatermarkOptions options)
        {
            _options = options;
        }

        public override void OnEndPage(PdfWriter writer, PdfDocument document)
        {
            if (_options == null || string.IsNullOrEmpty(_options.Text))
                return;

            var cb = writer.DirectContent;
            var pageSize = document.PageSize;
            float x = pageSize.Width / 2f;
            float y = pageSize.Height / 2f;
            cb.SaveState();
            cb.SetColorFill(BaseColor.Gray);
            cb.BeginText();
            cb.SetFontAndSize("F1", _options.FontSize > 0 ? _options.FontSize : 60);
            cb.SetTextMatrix(1, 0, 0, 1, x, y);
            cb.ShowText(_options.Text);
            cb.EndText();
            cb.RestoreState();
        }
    }