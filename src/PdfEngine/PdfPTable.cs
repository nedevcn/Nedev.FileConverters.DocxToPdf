namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// PDF表格单元格
/// </summary>
public class PdfPCell : IElement
{
    public int Type => 20;
    public bool IsContent() => true;
    public bool IsNestable() => true;

    public int Colspan { get; set; } = 1;
    public int Rowspan { get; set; } = 1;
    public int ColIndex { get; set; } = 0;

    public int HorizontalAlignment { get; set; } = Element.ALIGN_LEFT;
    public int VerticalAlignment { get; set; } = Element.ALIGN_MIDDLE;

    public float PaddingTop { get; set; } = 2f;
    public float PaddingBottom { get; set; } = 2f;
    public float PaddingLeft { get; set; } = 5.4f;
    public float PaddingRight { get; set; } = 5.4f;

    public float BorderWidthTop { get; set; } = 0.5f;
    public float BorderWidthBottom { get; set; } = 0.5f;
    public float BorderWidthLeft { get; set; } = 0.5f;
    public float BorderWidthRight { get; set; } = 0.5f;

    public BaseColor? BorderColorTop { get; set; }
    public BaseColor? BorderColorBottom { get; set; }
    public BaseColor? BorderColorLeft { get; set; }
    public BaseColor? BorderColorRight { get; set; }

    public BaseColor? BackgroundColor { get; set; }

    public float MinimumHeight { get; set; }
    public float FixedHeight { get; set; }

    public bool UseAscender { get; set; }
    public bool UseDescender { get; set; }
    public bool UseVariableBorders { get; set; }

    public int Border
    {
        get => _border;
        set
        {
            _border = value;
            var width = value == NO_BORDER ? 0f : 0.5f;
            BorderWidthTop = width;
            BorderWidthBottom = width;
            BorderWidthLeft = width;
            BorderWidthRight = width;
        }
    }
    private int _border = 1;

    // 兼容性属性
    public float Padding
    {
        get => PaddingTop;
        set { PaddingTop = value; PaddingBottom = value; PaddingLeft = value; PaddingRight = value; }
    }

    public float BorderWidth
    {
        get => BorderWidthTop;
        set { BorderWidthTop = value; BorderWidthBottom = value; BorderWidthLeft = value; BorderWidthRight = value; }
    }

    public BaseColor? BorderColor
    {
        get => BorderColorTop;
        set { BorderColorTop = value; BorderColorBottom = value; BorderColorLeft = value; BorderColorRight = value; }
    }

    private readonly List<IElement> _elements = [];

    public const int NO_BORDER = 0;

    public PdfPCell() { }

    public PdfPCell(Phrase phrase)
    {
        AddElement(phrase);
    }

    public PdfPCell(Paragraph paragraph)
    {
        AddElement(paragraph);
    }

    public void AddElement(IElement element)
    {
        _elements.Add(element);
    }

    public IEnumerable<IElement> Elements => _elements;

    public Phrase? Phrase
    {
        set
        {
            _elements.Clear();
            if (value != null) _elements.Add(value);
        }
    }
}

/// <summary>
/// PDF表格行
/// </summary>
public class PdfPRow
{
    public List<PdfPCell> Cells { get; } = [];
    public float MaxHeights { get; set; }
}

/// <summary>
/// PDF表格
/// </summary>
public class PdfPTable : IElement
{
    public int Type => 21;
    public bool IsContent() => true;
    public bool IsNestable() => false;

    public int NumberOfColumns { get; }
    public float WidthPercentage { get; set; } = 100f;
    public float SpacingBefore { get; set; } = 6f;
    public float SpacingAfter { get; set; } = 6f;
    public int HeaderRows { get; set; }
    public bool KeepTogether { get; set; }

    public float BorderWidth { set => DefaultCell.Border = value > 0 ? 1 : 0; }
    public BaseColor BorderColor { set => DefaultCell.BorderColorTop = value; }

    public PdfPCell DefaultCell { get; } = new();

    // 兼容性属性
    public object? TableEvent { get; set; }
    public const int LINECANVAS = 0;

    public List<PdfPRow> RowsList { get; } = [];
    private readonly List<float> _widths = [];

    public PdfPTable(int numColumns)
    {
        if (numColumns <= 0) throw new ArgumentException("列数必须大于0", nameof(numColumns));
        NumberOfColumns = numColumns;
    }

    public void AddCell(PdfPCell cell)
    {
        if (RowsList.Count == 0) RowsList.Add(new PdfPRow());
        
        var lastRow = RowsList.Last();
        int currentCols = lastRow.Cells.Sum(c => c.Colspan);
        
        if (currentCols >= NumberOfColumns)
        {
            lastRow = new PdfPRow();
            RowsList.Add(lastRow);
            currentCols = 0;
        }
        
        cell.ColIndex = currentCols;
        lastRow.Cells.Add(cell);
    }

    public void AddCell(string text)
    {
        var defaultFont = DefaultCell.Elements.FirstOrDefault() switch
        {
            Chunk c => c.Font,
            Phrase p => p.Font,
            _ => null
        };
        var cell = new PdfPCell(new Phrase(text, defaultFont ?? Font.Helvetica(12)));
        CopyCellStyle(cell);
        AddCell(cell);
    }

    public void AddCell(Phrase phrase)
    {
        var cell = new PdfPCell(phrase);
        CopyCellStyle(cell);
        AddCell(cell);
    }

    private void CopyCellStyle(PdfPCell cell)
    {
        cell.BorderWidthTop = DefaultCell.BorderWidthTop;
        cell.BorderWidthBottom = DefaultCell.BorderWidthBottom;
        cell.BorderWidthLeft = DefaultCell.BorderWidthLeft;
        cell.BorderWidthRight = DefaultCell.BorderWidthRight;
        cell.BorderColorTop = DefaultCell.BorderColorTop;
        cell.BorderColorBottom = DefaultCell.BorderColorBottom;
        cell.BorderColorLeft = DefaultCell.BorderColorLeft;
        cell.BorderColorRight = DefaultCell.BorderColorRight;
        cell.PaddingTop = DefaultCell.PaddingTop;
        cell.PaddingBottom = DefaultCell.PaddingBottom;
        cell.PaddingLeft = DefaultCell.PaddingLeft;
        cell.PaddingRight = DefaultCell.PaddingRight;
        cell.HorizontalAlignment = DefaultCell.HorizontalAlignment;
        cell.VerticalAlignment = DefaultCell.VerticalAlignment;
        cell.BackgroundColor = DefaultCell.BackgroundColor;
    }

    public void SetWidths(float[] widths)
    {
        _widths.Clear();
        _widths.AddRange(widths);
    }

    public float[] GetWidths(float totalWidth)
    {
        if (_widths.Count == NumberOfColumns)
        {
            var sum = _widths.Sum();
            return _widths.Select(w => w / sum * totalWidth).ToArray();
        }

        // 默认均分
        var width = totalWidth / NumberOfColumns;
        return Enumerable.Repeat(width, NumberOfColumns).ToArray();
    }

    public IEnumerable<PdfPCell> Cells => RowsList.SelectMany(r => r.Cells);

    public int Rows => RowsList.Count;
}
