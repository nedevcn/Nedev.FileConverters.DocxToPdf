namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// PDF文档类
/// </summary>
public class PdfDocument : IDisposable
{
    private readonly List<PdfPage> _pages = [];
    private Rectangle _pageSize;
    private float _marginLeft;
    private float _marginRight;
    private float _marginTop;
    private float _marginBottom;

    public Rectangle PageSize => _pageSize;
    public float MarginLeft => _marginLeft;
    public float MarginRight => _marginRight;
    public float MarginTop => _marginTop;
    public float MarginBottom => _marginBottom;

    public bool IsOpen { get; private set; }
    public int PageNumber => _pages.Count;

    public PdfDocument(Rectangle? pageSize = null, float marginLeft = 72f, float marginRight = 72f, float marginTop = 72f, float marginBottom = 72f)
    {
        _pageSize = pageSize ?? Rectangle.A4;
        _marginLeft = marginLeft;
        _marginRight = marginRight;
        _marginTop = marginTop;
        _marginBottom = marginBottom;
    }

    public void Open()
    {
        IsOpen = true;
    }

    public void Close()
    {
        IsOpen = false;
    }

    public void SetPageSize(Rectangle pageSize)
    {
        _pageSize = pageSize;
    }

    public void SetMargins(float left, float right, float top, float bottom)
    {
        _marginLeft = left;
        _marginRight = right;
        _marginTop = top;
        _marginBottom = bottom;
    }

    public event EventHandler<PdfPage>? PageAdded;

    public PdfPage NewPage()
    {
        var page = new PdfPage(_pages.Count + 1, _pageSize, this);
        _pages.Add(page);
        PageAdded?.Invoke(this, page);
        return page;
    }

    public void Add(IElement element)
    {
        if (!IsOpen) throw new InvalidOperationException("文档未打开");

        if (_pages.Count == 0)
        {
            NewPage();
        }

        var currentPage = _pages[^1];
        currentPage.AddElement(element);
    }

    public IEnumerable<PdfPage> Pages => _pages;

    public void Dispose()
    {
        Close();
    }
}

/// <summary>
/// PDF页面
/// </summary>
public class PdfPage
{
    public int PageNumber { get; }
    public Rectangle PageSize { get; }
    public PdfDocument Document { get; }

    private readonly List<IElement> _elements = [];

    public PdfPage(int pageNumber, Rectangle pageSize, PdfDocument document)
    {
        PageNumber = pageNumber;
        PageSize = pageSize;
        Document = document;
    }

    public void AddElement(IElement element)
    {
        _elements.Add(element);
    }

    public IEnumerable<IElement> Elements => _elements;
}
