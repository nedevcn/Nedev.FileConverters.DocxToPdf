namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// 分栏文本排版
/// </summary>
public class ColumnText
{
    private readonly PdfContentByte _canvas;
    private readonly List<IElement> _elements = [];
    private float _llx, _lly, _urx, _ury;
    private float _yLine;
    private int _currentPageNumber = 1;
    private AnnotationCollection? _annotations;

    public float YLine
    {
        get => _yLine;
        set => _yLine = value;
    }

    public const int NO_MORE_COLUMN = 1;
    public const int NO_MORE_TEXT = 2;

    public ColumnText(PdfContentByte canvas)
    {
        _canvas = canvas;
    }

    public void SetAnnotationCollection(AnnotationCollection annotations)
    {
        _annotations = annotations;
    }

    public void SetCurrentPage(int pageNumber)
    {
        _currentPageNumber = pageNumber;
    }

    public void SetSimpleColumn(float llx, float lly, float urx, float ury)
    {
        _llx = llx;
        _lly = lly;
        _urx = urx;
        _ury = ury;
        _yLine = ury;
    }

    public void AddElement(IElement element)
    {
        _elements.Add(element);
    }

    public int Go(bool simulate = false)
    {
        var remaining = new List<IElement>();
        var hasMoreText = false;

        foreach (var element in _elements)
        {
            if (_yLine <= _lly)
            {
                remaining.Add(element);
                hasMoreText = true;
                continue;
            }

            if (!simulate)
            {
                _yLine = RenderElement(element, _llx, _yLine, _urx);
            }
            else
            {
                _yLine -= EstimateHeight(element, _urx - _llx);
            }
        }

        _elements.Clear();
        _elements.AddRange(remaining);

        if (hasMoreText) return NO_MORE_COLUMN;
        return NO_MORE_TEXT;
    }

    private float RenderElement(IElement element, float x, float y, float maxX, bool simulate = false)
    {
        var width = maxX - x;

        switch (element)
        {
            case Paragraph para:
                return RenderParagraph(para, x, y, width, simulate);
            case Chunk chunk:
                return RenderChunk(chunk, x, y, simulate);
            case Phrase phrase:
                return RenderPhrase(phrase, x, y, simulate);
            case PdfPTable table:
                return RenderTable(table, x, y, width, simulate);
            case List list:
                return RenderList(list, x, y, width, simulate);
            case Image img:
                if (!simulate)
                {
                    _canvas.SaveState();
                    _canvas.AddImage(img, x, y - img.ScaledHeight);
                    _canvas.RestoreState();
                }
                return y - img.ScaledHeight - 4f; // Add a small margin
            default:
                if (element.Type == -100) // FloatingObject
                {
                    var floatObj = element as global::Nedev.FileConverters.DocxToPdf.Converters.FloatingObject;
                    if (floatObj != null && !simulate)
                    {
                        var imgObj = floatObj.Image;
                        _canvas.SaveState();
                        _canvas.AddImage(imgObj, floatObj.Left, floatObj.PositionIsAbsolute ? floatObj.Top : y - floatObj.Top - imgObj.ScaledHeight);
                        _canvas.RestoreState();
                    }
                    return floatObj != null && floatObj.Wrapping == global::Nedev.FileConverters.DocxToPdf.Converters.WrappingStyle.Inline ? y - floatObj.Height : y;
                }
                return y;
        }
    }

    private float RenderParagraph(Paragraph para, float x, float y, float width, bool simulate = false)
    {
        var lineHeight = para.Leading + para.Font.Size * para.MultipliedLeading;
        if (lineHeight <= 0) lineHeight = para.Font?.Size * para.MultipliedLeading ?? 16f; // Fallback

        y -= para.SpacingBefore;

        if (!simulate && para.RenderedCallback != null)
        {
            para.RenderedCallback(para, _currentPageNumber);
        }

        // calculate the usable width inside paragraph indentations
        float availableWidth = width - para.IndentationLeft - para.IndentationRight;
        if (availableWidth < 0) availableWidth = 0;

        // break chunks into lines first so we can compute each line's width
        var lines = new List<(List<Chunk> chunks, float lineWidth)>();
        var currentLine = new List<Chunk>();
        float currentLineWidth = 0;
        bool firstChunkOnLine = true;

        foreach (var chunk in para.Chunks)
        {
            var cw = chunk.GetWidth();
            if (!firstChunkOnLine && currentLineWidth + cw > availableWidth && currentLine.Count > 0)
            {
                lines.Add((currentLine, currentLineWidth));
                currentLine = new List<Chunk>();
                currentLineWidth = 0;
                firstChunkOnLine = true;
            }

            currentLine.Add(chunk);
            currentLineWidth += cw;
            firstChunkOnLine = false;
        }
        if (currentLine.Count > 0)
        {
            lines.Add((currentLine, currentLineWidth));
        }

        // render each line applying alignment and indentation
        bool firstLine = true;
        foreach (var (chunks, lineWidth) in lines)
        {
            float startX = x + para.IndentationLeft;
            if (firstLine)
            {
                startX += para.FirstLineIndent;
                firstLine = false;
            }

            if (para.Alignment == Element.ALIGN_CENTER)
            {
                startX += Math.Max(0, (availableWidth - lineWidth) / 2f);
            }
            else if (para.Alignment == Element.ALIGN_RIGHT)
            {
                startX += Math.Max(0, availableWidth - lineWidth);
            }

            var currentX = startX;
            foreach (var chunk in chunks)
            {
                currentX = RenderChunk(chunk, currentX, y, simulate);
            }

            y -= lineHeight;
        }

        return y - para.SpacingAfter;
    }

    private float RenderChunk(Chunk chunk, float x, float y, bool simulate = false)
    {
        if (string.IsNullOrEmpty(chunk.Content)) return y;

        if (!simulate)
        {
            _canvas.SaveState();

            if (chunk.BackgroundColor != null)
            {
                _canvas.SetColorFill(chunk.BackgroundColor);
                _canvas.Rectangle(x, y - chunk.Font.Size * 0.2f, chunk.GetWidth(), chunk.Font.Size * 1.2f);
                _canvas.Fill();
            }

            _canvas.SetColorFill(chunk.Font.Color);
            var textBaselineY = y - chunk.Font.Size * 0.8f + chunk.TextRise;

            _canvas.BeginText();
            _canvas.SetFontAndSize("F1", chunk.Font.Size);
            _canvas.SetTextMatrix(1, 0, 0, 1, x, textBaselineY);
            _canvas.ShowText(chunk.Content);
            _canvas.EndText();

            if (chunk.HasUnderline)
            {
                _canvas.SetLineWidth(chunk.UnderlineThickness);
                _canvas.SetColorStroke(chunk.Font.Color);
                _canvas.MoveTo(x, textBaselineY + chunk.UnderlineYPosition);
                _canvas.LineTo(x + chunk.GetWidth(), textBaselineY + chunk.UnderlineYPosition);
                _canvas.Stroke();
            }

            if (_annotations != null && !string.IsNullOrEmpty(chunk.Anchor))
            {
                var chunkHeight = chunk.Font.Size * 1.2f;
                _annotations.AddLink(_currentPageNumber, x, y - chunkHeight, chunk.GetWidth(), chunkHeight, chunk.Anchor);
            }

            _canvas.RestoreState();
        }

        return y;
    }

    private float RenderPhrase(Phrase phrase, float x, float y, bool simulate = false)
    {
        var currentX = x;
        foreach (var chunk in phrase.Chunks)
        {
            RenderChunk(chunk, currentX, y, simulate);
            currentX += chunk.GetWidth();
        }
        return y - phrase.Font.Size;
    }

    private float RenderTable(PdfPTable table, float x, float y, float width, bool simulate = false)
    {
        y -= table.SpacingBefore;

        var widths = table.GetWidths(width);

        foreach (var row in table.RowsList)
        {
            var cellsInRow = row.Cells;

            if (cellsInRow.Count == 0) continue;

            // 先计算该行所需的最大高度
            float maxRowHeight = cellsInRow[0].Elements.FirstOrDefault() is Paragraph p ? p.Font.Size * p.MultipliedLeading : 20f;
            for (int c = 0; c < cellsInRow.Count; c++)
            {
                var cell = cellsInRow[c];
                var colIndex = cell.ColIndex;
                var colSpan = cell.Colspan;
                
                float cellColWidth = 0;
                for (int s = 0; s < colSpan && (colIndex + s) < widths.Length; s++) {
                    cellColWidth += widths[colIndex + s] * table.WidthPercentage / 100f;
                }

                float cellHeight = cell.PaddingTop + cell.PaddingBottom;
                float simY = 0;
                foreach (var elem in cell.Elements)
                {
                    var availableWidth = cellColWidth - cell.PaddingLeft - cell.PaddingRight;
                    if (availableWidth < 1) availableWidth = 1;
                    simY = RenderElement(elem, 0, simY, availableWidth, simulate: true);
                }
                cellHeight += -simY;
                if (cellHeight > maxRowHeight) maxRowHeight = cellHeight;
            }

            // 进行实际排版（模拟或绘制）
            for (int c = 0; c < cellsInRow.Count; c++)
            {
                var cellToDraw = cellsInRow[c];
                var colIndex = cellToDraw.ColIndex;
                var colSpan = cellToDraw.Colspan;
                
                var currentX = x;
                // 计算当前单元格的起始 X 坐标
                for (int s = 0; s < colIndex && s < widths.Length; s++) {
                    currentX += widths[s] * table.WidthPercentage / 100f;
                }
                
                float cellWidth = 0;
                for (int s = 0; s < colSpan && (colIndex + s) < widths.Length; s++) {
                    cellWidth += widths[colIndex + s] * table.WidthPercentage / 100f;
                }

                if (!simulate)
                {
                    if (cellToDraw.BackgroundColor != null)
                    {
                        _canvas.SaveState();
                        _canvas.SetColorFill(cellToDraw.BackgroundColor);
                        _canvas.Rectangle(currentX, y - maxRowHeight, cellWidth, maxRowHeight);
                        _canvas.Fill();
                        _canvas.RestoreState();
                    }

                    // 绘制边框
                    if (cellToDraw.BorderWidthTop > 0)
                    {
                        _canvas.SaveState();
                        _canvas.SetLineWidth(cellToDraw.BorderWidthTop);
                        _canvas.SetColorStroke(cellToDraw.BorderColorTop ?? BaseColor.Black);
                        _canvas.MoveTo(currentX, y);
                        _canvas.LineTo(currentX + cellWidth, y);
                        _canvas.Stroke();
                        _canvas.RestoreState();
                    }
                    if (cellToDraw.BorderWidthBottom > 0)
                    {
                        _canvas.SaveState();
                        _canvas.SetLineWidth(cellToDraw.BorderWidthBottom);
                        _canvas.SetColorStroke(cellToDraw.BorderColorBottom ?? BaseColor.Black);
                        _canvas.MoveTo(currentX, y - maxRowHeight);
                        _canvas.LineTo(currentX + cellWidth, y - maxRowHeight);
                        _canvas.Stroke();
                        _canvas.RestoreState();
                    }
                    if (cellToDraw.BorderWidthLeft > 0)
                    {
                        _canvas.SaveState();
                        _canvas.SetLineWidth(cellToDraw.BorderWidthLeft);
                        _canvas.SetColorStroke(cellToDraw.BorderColorLeft ?? BaseColor.Black);
                        _canvas.MoveTo(currentX, y);
                        _canvas.LineTo(currentX, y - maxRowHeight);
                        _canvas.Stroke();
                        _canvas.RestoreState();
                    }
                    if (cellToDraw.BorderWidthRight > 0)
                    {
                        _canvas.SaveState();
                        _canvas.SetLineWidth(cellToDraw.BorderWidthRight);
                        _canvas.SetColorStroke(cellToDraw.BorderColorRight ?? BaseColor.Black);
                        _canvas.MoveTo(currentX + cellWidth, y);
                        _canvas.LineTo(currentX + cellWidth, y - maxRowHeight);
                        _canvas.Stroke();
                        _canvas.RestoreState();
                    }
                }

                var contentY = y - cellToDraw.PaddingTop;
                foreach (var elem in cellToDraw.Elements)
                {
                    var availableWidth = cellWidth - cellToDraw.PaddingLeft - cellToDraw.PaddingRight;
                    if (availableWidth < 1) availableWidth = 1;
                    contentY = RenderElement(elem, currentX + cellToDraw.PaddingLeft, contentY, currentX + cellToDraw.PaddingLeft + availableWidth, simulate);
                }
            }

            y -= maxRowHeight;
        }

        return y - table.SpacingAfter;
    }

    private float RenderList(List list, float x, float y, float width, bool simulate = false)
    {
        var itemNumber = 1;

        foreach (var item in list.Items)
        {
            var symbol = list.ListType == List.ORDERED ? $"{itemNumber}." : list.ListSymbol.Content;
            var symbolChunk = new Chunk(symbol, item.Font);
            RenderChunk(symbolChunk, x + list.IndentationLeft - list.SymbolIndent, y, simulate);

            y = RenderParagraph(item, x + list.IndentationLeft, y, width - list.IndentationLeft, simulate);
            itemNumber++;
        }

        return y;
    }

    private float EstimateHeight(IElement element, float width)
    {
        // 直接使用 simulate = true 来精确计算耗费高度
        return -RenderElement(element, 0, 0, width, simulate: true);
    }

    public static bool HasMoreText(int status) => status == NO_MORE_COLUMN;
}
