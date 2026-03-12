using Nedev.FileConverters.DocxToPdf.Models;

namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// 分栏文本排版
/// </summary>
public class ColumnText
{
    private readonly PdfContentByte _canvas;
    private readonly List<IElement> _elements = [];

    /// <summary>
    /// Public view of the pending element queue (used by unit tests).
    /// </summary>
    public IReadOnlyList<IElement> Elements => _elements.AsReadOnly();
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

    // 环绕排除区
    public List<SkiaSharp.SKRect> Exclusions { get; } = new();

    // 行号相关
    public LineNumberSettings? LineNumberSettings { get; set; }
    public int CurrentLineNumber { get; set; } = 1;
    private int _lastPageNumber = 0;

    // 文本方向
    public TextDirection TextDirection { get; set; } = TextDirection.Horizontal;

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
        if (_currentPageNumber != pageNumber)
        {
            _currentPageNumber = pageNumber;
            // 处理每页重置行号
            if (LineNumberSettings?.RestartMode == LineNumberRestartMode.NewPage)
            {
                CurrentLineNumber = LineNumberSettings.Start;
            }
        }
    }

    public void SetSimpleColumn(float llx, float lly, float urx, float ury)
    {
        _llx = llx;
        _lly = lly;
        _urx = urx;
        _ury = ury;
        if (TextDirection == TextDirection.Vertical)
            _yLine = urx;
        else
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
            int startCount = _elements.Count;

            foreach (var element in _elements)
            {
                bool boundaryHit = false;
                if (TextDirection == TextDirection.Vertical)
                {
                    if (_yLine <= _llx) // 左边界
                    {
                        boundaryHit = true;
                    }
                }
                else
                {
                    if (_yLine <= _lly) // 下边界
                    {
                        boundaryHit = true;
                    }
                }

                if (boundaryHit)
                {
                    // element does not fit in current column/page
                    // attempt to split paragraph across the boundary
                    if (element is Paragraph para)
                    {
                        // calculate how much vertical space is left
                        float availHeight = TextDirection == TextDirection.Vertical
                            ? _yLine - _llx
                            : _yLine - _lly;
                        float availLength = TextDirection == TextDirection.Vertical
                            ? _ury - _lly
                            : _urx - _llx;

                        var (first, rest) = SplitParagraphByHeight(para, availLength, availHeight);
                        if (first != null && rest != null)
                        {
                            // render the portion that fits
                            if (!simulate)
                            {
                                if (TextDirection == TextDirection.Vertical)
                                    _yLine = RenderElement(first, _ury, _yLine, _lly);
                                else
                                    _yLine = RenderElement(first, _llx, _yLine, _urx);
                            }
                            else
                            {
                                _yLine -= EstimateHeight(first, availLength);
                            }
                            remaining.Add(rest);
                            hasMoreText = true;
                            continue;
                        }
                        // if no sub‑paragraph fit, fall through and treat as whole element
                    }

                    remaining.Add(element);
                    hasMoreText = true;
                    continue;
                }

                if (!simulate)
                {
                    if (TextDirection == TextDirection.Vertical)
                    {
                        _yLine = RenderElement(element, _ury, _yLine, _lly);
                    }
                    else
                    {
                        _yLine = RenderElement(element, _llx, _yLine, _urx);
                    }
                }
                else
                {
                    if (TextDirection == TextDirection.Vertical)
                    {
                        _yLine -= EstimateHeight(element, _ury - _lly);
                    }
                    else
                    {
                        _yLine -= EstimateHeight(element, _urx - _llx);
                    }
                }
            }

            _elements.Clear();
            _elements.AddRange(remaining);

            if (hasMoreText)
            {
                // if nothing was rendered at all (remaining == original list) then give up to avoid infinite loops
                if (remaining.Count == startCount)
                    return NO_MORE_TEXT;
                return NO_MORE_COLUMN;
            }
    /// 渲染元素
    /// </summary>
    /// <param name="element">元素</param>
    /// <param name="startInline">内联起始位置 (Horizontal: Left, Vertical: Top)</param>
    /// <param name="startBlock">块起始位置 (Horizontal: Top, Vertical: Right)</param>
    /// <param name="limitInline">内联限制位置 (Horizontal: Right, Vertical: Bottom)</param>
    /// <param name="simulate">是否模拟</param>
    /// <returns>新的块位置 (Horizontal: newY, Vertical: newX)</returns>
    private float RenderElement(IElement element, float startInline, float startBlock, float limitInline, bool simulate = false)
    {
        var availableLength = TextDirection == TextDirection.Vertical 
            ? startInline - limitInline // Top - Bottom
            : limitInline - startInline; // Right - Left

        switch (element)
        {
            case Paragraph para:
                return RenderParagraph(para, startInline, startBlock, availableLength, simulate);
            case Chunk chunk:
                if (TextDirection == TextDirection.Vertical)
                {
                    RenderChunkVertical(chunk, startBlock, startInline, simulate);
                    return startBlock;
                }
                return RenderChunk(chunk, startInline, startBlock, simulate);
            case Phrase phrase:
                return RenderParagraph(new Paragraph(phrase), startInline, startBlock, availableLength, simulate);
            case PdfPTable table:
                return RenderTable(table, startInline, startBlock, availableLength, simulate);
            case List list:
                return RenderList(list, startInline, startBlock, availableLength, simulate);
            case Image img:
                if (!simulate)
                {
                    _canvas.SaveState();
                    if (TextDirection == TextDirection.Vertical)
                        _canvas.AddImage(img, startBlock - img.ScaledWidth, startInline - img.ScaledHeight);
                    else
                        _canvas.AddImage(img, startInline, startBlock - img.ScaledHeight);
                    _canvas.RestoreState();
                }
                return startBlock - (TextDirection == TextDirection.Vertical ? img.ScaledWidth : img.ScaledHeight) - 4f; 
            default:
                if (element.Type == -100) // FloatingObject
                {
                    var floatObj = element as global::Nedev.FileConverters.DocxToPdf.Converters.FloatingObject;
                    if (floatObj != null && !simulate)
                    {
                        var imgObj = floatObj.Image;
                        _canvas.SaveState();
                        if (TextDirection == TextDirection.Vertical && !floatObj.PositionIsAbsolute)
                        {
                            // 竖排：startBlock是X，startInline是Y
                            _canvas.AddImage(imgObj, startBlock - floatObj.Width - floatObj.Left, startInline - floatObj.Top - imgObj.ScaledHeight);
                        }
                        else
                        {
                            _canvas.AddImage(imgObj, floatObj.Left, floatObj.PositionIsAbsolute ? floatObj.Top : startBlock - floatObj.Top - imgObj.ScaledHeight);
                        }
                        _canvas.RestoreState();
                    }
                    return floatObj != null && floatObj.Wrapping == global::Nedev.FileConverters.DocxToPdf.Converters.WrappingStyle.Inline 
                        ? startBlock - (TextDirection == TextDirection.Vertical ? floatObj.Width : floatObj.Height) 
                        : startBlock;
                }
                return startBlock;
        }
    }

    private float RenderParagraph(Paragraph para, float startInline, float startBlock, float availableLength, bool simulate = false)
    {
        if (TextDirection == TextDirection.Vertical)
        {
            return RenderParagraphVertical(para, startInline, startBlock, availableLength, simulate);
        }

        var x = startInline;
        var y = startBlock;
        var width = availableLength;
        // ... (rest of RenderParagraph horizontal logic)
        var lineHeight = para.Leading + para.Font.Size * para.MultipliedLeading;
        if (lineHeight <= 0) lineHeight = para.Font?.Size * para.MultipliedLeading ?? 16f; // Fallback

        y -= para.SpacingBefore;

        if (!simulate && para.RenderedCallback != null)
        {
            para.RenderedCallback(para, _currentPageNumber);
        }

        // calculate the usable width inside paragraph indentations
        float defaultAvailableWidth = width - para.IndentationLeft - para.IndentationRight;
        if (defaultAvailableWidth < 0) defaultAvailableWidth = 0;

        // break chunks into lines using Y-aware exclusion wrapping
        var lines = new List<(List<Chunk> chunks, float lineWidth, float lineStartX, float yLine, float lineAvailWidth)>();
        var chunksList = new List<Chunk>(para.Chunks);
        int chunkIdx = 0;

        float currentY = y - para.SpacingBefore;
        
        var currentLine = new List<Chunk>();
        float currentLineWidth = 0;
        bool firstTokenOnLine = true;
        
        float _currentLineStartX = x + para.IndentationLeft + para.FirstLineIndent;
        float _currentLineAvailWidth = defaultAvailableWidth;
        bool _needEval = true;

        while (chunkIdx < chunksList.Count)
        {
            if (_needEval)
            {
                _currentLineStartX = x + para.IndentationLeft + (lines.Count == 0 ? para.FirstLineIndent : 0);
                _currentLineAvailWidth = defaultAvailableWidth;
                
                // ex.Bottom is smaller Y, ex.Top is larger Y. PDF Y goes up.
                foreach (var ex in Exclusions)
                {
                    if (currentY - lineHeight < ex.Top && currentY > ex.Bottom)
                    {
                        if (ex.Left <= _currentLineStartX + _currentLineAvailWidth / 2) // image on left half
                        {
                            float shift = Math.Max(0, (ex.Right + 8f) - _currentLineStartX); // 8f padding
                            _currentLineStartX += shift;
                            _currentLineAvailWidth -= shift;
                        }
                        else // image on right half
                        {
                            float cut = Math.Max(0, (_currentLineStartX + _currentLineAvailWidth) - (ex.Left - 8f));
                            _currentLineAvailWidth -= cut;
                        }
                    }
                }
                
                if (_currentLineAvailWidth <= 20 && currentY > _lly)
                {
                    currentY -= lineHeight; // move down and try again
                    continue;
                }
                _needEval = false;
            }

            var chunk = chunksList[chunkIdx];
            var cw = chunk.GetWidth();

            if (firstTokenOnLine || currentLineWidth + cw <= _currentLineAvailWidth)
            {
                currentLine.Add(chunk);
                currentLineWidth += cw;
                firstTokenOnLine = false;
                chunkIdx++;
                continue;
            }

            // Chunk overflows — split by words
            chunksList.RemoveAt(chunkIdx);
            var subChunks = SplitChunkByWords(chunk);
            chunksList.InsertRange(chunkIdx, subChunks);
            
            var sw = chunksList[chunkIdx].GetWidth();
            
            // If it still overflows after splitting, break the line
            if (firstTokenOnLine || currentLineWidth + sw <= _currentLineAvailWidth)
            {
                currentLine.Add(chunksList[chunkIdx]);
                currentLineWidth += sw;
                firstTokenOnLine = false;
                chunkIdx++;
            }
            else
            {
                lines.Add((currentLine, currentLineWidth, _currentLineStartX, currentY, _currentLineAvailWidth));
                currentLine = new List<Chunk>();
                currentLineWidth = 0;
                firstTokenOnLine = true;
                currentY -= lineHeight;
                _needEval = true;
            }
        }
        
        if (currentLine.Count > 0)
        {
            lines.Add((currentLine, currentLineWidth, _currentLineStartX, currentY, _currentLineAvailWidth));
        }

        // render each line applying alignment and indentation
        foreach (var (chunks, lineWidth, lineStartX, yLine, lineAvailWidth) in lines)
        {
            float startX = lineStartX;

            if (para.Alignment == Element.ALIGN_CENTER)
            {
                startX += Math.Max(0, (lineAvailWidth - lineWidth) / 2f);
            }
            else if (para.Alignment == Element.ALIGN_RIGHT)
            {
                startX += Math.Max(0, lineAvailWidth - lineWidth);
            }

            var currentX = startX;
            foreach (var chunk in chunks)
            {
                currentX = RenderChunk(chunk, currentX, yLine, simulate);
            }

            // 绘制行号
            if (!simulate && LineNumberSettings != null)
            {
                if (CurrentLineNumber % LineNumberSettings.CountBy == 0)
                {
                    float lnX = _llx - LineNumberSettings.Distance;
                    // 行号通常右对齐到距离位置？Word是右对齐到 margin - distance。
                    // 简单实现：左对齐或右对齐
                    // 这里假设 LineNumberSettings.Distance 是距离正文的间距
                    // 那么行号应该画在 _llx - Distance - (行号宽度) 
                    // 简化：画在 _llx - Distance 处，右对齐
                    
                    _canvas.SaveState();
                    _canvas.BeginText();
                    // 使用 Helvetica 作为行号字体，确保可用性
                    var lineNumberFontName = FontFactory.IsRegistered("Helvetica") ? "Helvetica" : "F1";
                    _canvas.SetFontAndSize(lineNumberFontName, para.Font.Size);
                    
                    var lnText = CurrentLineNumber.ToString();
                    // 简单估算宽度
                    float lnWidth = lnText.Length * para.Font.Size * 0.5f; 
                    
                    _canvas.SetTextMatrix(1, 0, 0, 1, lnX - lnWidth, yLine - para.Font.Size * 0.8f);
                    _canvas.ShowText(lnText);
                    _canvas.EndText();
                    _canvas.RestoreState();
                }
                CurrentLineNumber++;
            }
        }
        
        if (lines.Count > 0)
            y = lines[^1].yLine - lineHeight;
        else
            y = currentY - lineHeight;

        foreach (var extra in para.ExtraElements)
        {
            y = RenderElement(extra, startInline, y, availableLength, simulate);
        }

        return y - para.SpacingAfter;
    }

    private static List<Chunk> SplitChunkByWords(Chunk chunk)
    {
        var result = new List<Chunk>();
        if (string.IsNullOrEmpty(chunk.Content))
        {
            result.Add(chunk);
            return result;
        }

        var words = chunk.Content.Split(new[] { ' ' }, StringSplitOptions.None);
        for (int i = 0; i < words.Length; i++)
        {
            string w = words[i];
            if (i < words.Length - 1)
                w += " ";
                
            if (!string.IsNullOrEmpty(w))
            {
                var newChunk = new Chunk(w, chunk.Font)
                {
                    BackgroundColor = chunk.BackgroundColor,
                    TextRise = chunk.TextRise,
                    Anchor = chunk.Anchor,
                    HasUnderline = chunk.HasUnderline,
                    UnderlineThickness = chunk.UnderlineThickness,
                    UnderlineYPosition = chunk.UnderlineYPosition
                };
                result.Add(newChunk);
            }
        }
        return result;
    }

    private float RenderParagraphVertical(Paragraph para, float topY, float rightX, float height, bool simulate = false)
    {
        // 竖排：startInline 是 topY，startBlock 是 rightX，availableLength 是 height
        // 字符从上到下排列，行从右到左排列
        
        // 竖排行高 = 字符宽度（或字号）+ 行距
        // 简单起见，假设竖排时行高由字号决定（横向宽度）
        var lineWidth = para.Leading + para.Font.Size * para.MultipliedLeading;
        if (lineWidth <= 0) lineWidth = para.Font?.Size * para.MultipliedLeading ?? 16f;

        // 右边距（段前距）
        rightX -= para.SpacingBefore;

        // 计算可用高度（减去缩进）
        // IndentationLeft 在竖排中对应 Top 缩进？还是 Right 缩进？
        // 通常 IndentationLeft 是段落起始边的缩进。
        // 竖排起始边是 Top。所以 IndentationLeft -> Top Indent, IndentationRight -> Bottom Indent.
        float availableHeight = height - para.IndentationLeft - para.IndentationRight;
        if (availableHeight < 0) availableHeight = 0;

        // 分行逻辑
        var lines = new List<(List<Chunk> chunks, float lineLength)>();
        var currentLine = new List<Chunk>();
        float currentLineLength = 0;
        bool firstChunkOnLine = true;

        foreach (var chunk in para.Chunks)
        {
            // 竖排字符高度：
            // CJK: 字号
            // Latin: 旋转后宽度 -> 字号
            // 简单假设所有字符高度 = font.Size
            // 如果内容包含字符串，需要逐字符计算高度
            // 这里简化：Chunk.GetWidth() 返回的是水平宽度。对于等宽中文字体，宽度=高度。
            // 对于非等宽字体，高度通常是固定的（字号）。
            // 我们假设 chunk 的“竖向长度”等于 chunk.Content.Length * chunk.Font.Size (简单估算)
            // 或者更准确：调用 GetWidthPoint() 但假设它是竖向的？
            // 更好的方法：RenderChunkVertical 负责绘制。这里只负责分行。
            // 假设每个字符高度 = Font.Size。
            
            float chunkHeight = 0;
            if (!string.IsNullOrEmpty(chunk.Content))
            {
                // 简单估算：每个字符高度 = Font.Size
                // 实际应区分半角全角
                // 这里暂时用 GetWidth() 近似，因为 CJK 宽度=高度。
                chunkHeight = chunk.GetWidth(); 
            }

            if (!firstChunkOnLine && currentLineLength + chunkHeight > availableHeight && currentLine.Count > 0)
            {
                lines.Add((currentLine, currentLineLength));
                currentLine = new List<Chunk>();
                currentLineLength = 0;
                firstChunkOnLine = true;
            }

            currentLine.Add(chunk);
            currentLineLength += chunkHeight;
            firstChunkOnLine = false;
        }
        if (currentLine.Count > 0)
        {
            lines.Add((currentLine, currentLineLength));
        }

        // 渲染每一行
        bool firstLine = true;
        foreach (var (chunks, lineLen) in lines)
        {
            // 计算当前行的起始 Y (Top)
            // IndentationLeft -> Top Indent
            float startY = topY - para.IndentationLeft;
            if (firstLine)
            {
                startY -= para.FirstLineIndent;
                firstLine = false;
            }

            // 对齐处理
            if (para.Alignment == Element.ALIGN_CENTER)
            {
                startY -= Math.Max(0, (availableHeight - lineLen) / 2f);
            }
            else if (para.Alignment == Element.ALIGN_RIGHT) // Bottom Align in Vertical
            {
                startY -= Math.Max(0, availableHeight - lineLen);
            }

            var currentY = startY;
            
            // 绘制当前行字符
            // 字符中心 X 对齐到 rightX - lineWidth/2 ? 
            // 或者 rightX 是行的右边界。字符画在 rightX - lineWidth 到 rightX 之间。
            // 字符中心 X = rightX - lineWidth / 2。
            // 简化：字符画在 rightX - Font.Size (假设 lineWidth ~ Font.Size)
            
            foreach (var chunk in chunks)
            {
                currentY = RenderChunkVertical(chunk, rightX, currentY, simulate);
            }

            // 绘制行号 (竖排时绘制在上方)
            if (!simulate && LineNumberSettings != null)
            {
                if (CurrentLineNumber % LineNumberSettings.CountBy == 0)
                {
                    float lnY = topY + LineNumberSettings.Distance + para.Font.Size;
                    _canvas.SaveState();
                    _canvas.BeginText();
                    // 使用 Helvetica 作为行号字体，确保可用性
                    var lineNumberFontName = FontFactory.IsRegistered("Helvetica") ? "Helvetica" : "F1";
                    _canvas.SetFontAndSize(lineNumberFontName, para.Font.Size);
                    
                    var lnText = CurrentLineNumber.ToString();
                    float lnWidth = lnText.Length * para.Font.Size * 0.5f; 
                    
                    _canvas.SetTextMatrix(1, 0, 0, 1, rightX - lineWidth / 2f - lnWidth / 2f, lnY);
                    _canvas.ShowText(lnText);
                    _canvas.EndText();
                    _canvas.RestoreState();
                }
                CurrentLineNumber++;
            }

            // 推进到下一行（向左）
            rightX -= lineWidth;
        }

        return rightX - para.SpacingAfter;
    }

    private float RenderChunk(Chunk chunk, float startInline, float startBlock, bool simulate = false)
    {
        if (TextDirection == TextDirection.Vertical)
        {
            return RenderChunkVertical(chunk, startBlock, startInline, simulate);
        }

        var x = startInline;
        var y = startBlock;

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
            // 使用 chunk 的字体族，确保在 PDF 中已注册
            var fontName = FontFactory.IsRegistered(chunk.Font.Family) ? chunk.Font.Family : "F1";
            _canvas.SetFontAndSize(fontName, chunk.Font.Size);
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

    private float RenderChunkVertical(Chunk chunk, float x, float y, bool simulate = false)
    {
        if (string.IsNullOrEmpty(chunk.Content))
            return y;

        // x is the right edge of the column; y is the top of the current character
        // characters flow from top to bottom, so y decreases after each glyph
        foreach (var c in chunk.Content)
        {
            float charHeight = chunk.Font.Size;   // vertical advance is always font size
            // use font metrics for width -- more accurate for Latin letters
            float charWidth = chunk.Font.GetWidthPoint(c.ToString());

            if (!simulate)
            {
                _canvas.SaveState();
                _canvas.BeginText();
                var fontName = FontFactory.IsRegistered(chunk.Font.Family) ? chunk.Font.Family : "F1";
                _canvas.SetFontAndSize(fontName, chunk.Font.Size);

                bool isCJK = (c >= '\u4e00' && c <= '\u9fff') ||
                             (c >= '\u3000' && c <= '\u303f') ||
                             (c >= '\uff00' && c <= '\uffef');

                if (isCJK)
                {
                    // draw normally, baseline slightly below top
                    float drawX = x - charWidth;
                    float drawY = y - charHeight * 0.8f;
                    _canvas.SetTextMatrix(1, 0, 0, 1, drawX, drawY);
                    _canvas.ShowText(c.ToString());
                }
                else
                {
                    // rotate clockwise 90° so the character's baseline runs downwards
                    float drawX = x - charWidth;
                    float drawY = y;
                    // transformation matrix [a b c d e f] where
                    // a=0, b=1, c=-1, d=0 produces a clockwise 90° rotation
                    _canvas.SetTextMatrix(0, 1, -1, 0, drawX, drawY);
                    _canvas.ShowText(c.ToString());
                }

                _canvas.EndText();
                _canvas.RestoreState();
            }

            y -= charHeight;
        }

        return y;
    }

    private float RenderPhrase(Phrase phrase, float startInline, float startBlock, bool simulate = false)
    {
        var currentInline = startInline;
        foreach (var chunk in phrase.Chunks)
        {
            if (TextDirection == TextDirection.Vertical)
            {
                // Vertical: startInline is Top, decreases (y)
                currentInline = RenderChunkVertical(chunk, startBlock, currentInline, simulate);
            }
            else
            {
                RenderChunk(chunk, currentInline, startBlock, simulate);
                currentInline += chunk.GetWidth();
            }
        }
        
        if (TextDirection == TextDirection.Vertical)
            return startBlock; // Phrase rendering doesn't advance block, paragraph does
        else
            return startBlock - phrase.Font.Size; // Approximate height for horizontal
    }

    private float RenderTable(PdfPTable table, float x, float y, float width, bool simulate = false)
    {
        y -= table.SpacingBefore;

        var widths = table.GetWidths(width);

        // Apply table horizontal alignment offset
        var tableActualWidth = widths.Sum() * table.WidthPercentage / 100f;
        if (table.HorizontalAlignment == Element.ALIGN_CENTER)
            x += Math.Max(0, (width - tableActualWidth) / 2f);
        else if (table.HorizontalAlignment == Element.ALIGN_RIGHT)
            x += Math.Max(0, width - tableActualWidth);

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

    /// <summary>
    /// Try to split a paragraph so that the first part fits in the given vertical
    /// <paramref name="availableHeight"/>.  The <paramref name="availableLength"/>
    /// parameter is the inline dimension (width for horizontal text, height for
    /// vertical).  Returns a tuple of (firstPart, remainder).  If splitting is not
    /// possible (nothing fits or paragraph already fits) the tuple elements will be
    /// null appropriately.
    /// </summary>
    private (Paragraph? first, Paragraph? remainder) SplitParagraphByHeight(Paragraph para, float availableLength, float availableHeight)
    {
        var chunks = para.Chunks;
        if (chunks.Count == 0)
            return (null, null);

        int fitCount = 0;
        for (int i = 1; i <= chunks.Count; i++)
        {
            var testPara = new Paragraph(para)
            {
                Chunks = chunks.Take(i).ToList()
            };
            float height = EstimateHeight(testPara, availableLength);
            if (height <= availableHeight + 0.1f)
            {
                fitCount = i;
            }
            else
            {
                break;
            }
        }

        if (fitCount == 0)
        {
            // not even a single chunk fits
            return (null, para);
        }

        if (fitCount >= chunks.Count)
        {
            // whole paragraph fits
            return (para, null);
        }

        var firstPara = new Paragraph(para)
        {
            Chunks = chunks.Take(fitCount).ToList()
        };
        var restPara = new Paragraph(para)
        {
            Chunks = chunks.Skip(fitCount).ToList()
        };
        return (firstPara, restPara);
    }

    public static bool HasMoreText(int status) => status == NO_MORE_COLUMN;
}
