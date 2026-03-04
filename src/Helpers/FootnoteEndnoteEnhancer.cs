using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Nedev.DocxToPdf.PdfEngine;
using iTextFont = Nedev.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.DocxToPdf.PdfEngine.Paragraph;
using iTextChunk = Nedev.DocxToPdf.PdfEngine.Chunk;

namespace Nedev.DocxToPdf.Helpers;

/// <summary>
/// 脚注与尾注增强处理器
/// </summary>
public static class FootnoteEndnoteEnhancer
{
    /// <summary>
    /// 脚注/尾注条目
    /// </summary>
    public class NoteEntry
    {
        public int Id { get; set; }
        public int Number { get; set; }
        public bool IsFootnote { get; set; }
        public string Content { get; set; } = "";
        public List<IElement> Elements { get; set; } = new();
        public string? ReferenceMark { get; set; }
    }

    /// <summary>
    /// 从脚注部分提取所有脚注
    /// </summary>
    public static List<NoteEntry> ExtractFootnotes(Footnotes? footnotes, Dictionary<int, int> numberMap)
    {
        var entries = new List<NoteEntry>();
        if (footnotes == null) return entries;

        foreach (var footnote in footnotes.Elements<Footnote>())
        {
            var idAttr = footnote.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
            if (string.IsNullOrEmpty(idAttr.Value) || !int.TryParse(idAttr.Value, out var id) || id <= 0) continue;

            var number = numberMap.TryGetValue(id, out var n) ? n : id;
            var entry = new NoteEntry
            {
                Id = id,
                Number = number,
                IsFootnote = true
            };

            // 提取内容
            ExtractNoteContent(footnote, entry);
            entries.Add(entry);
        }

        return entries;
    }

    /// <summary>
    /// 从尾注部分提取所有尾注
    /// </summary>
    public static List<NoteEntry> ExtractEndnotes(Endnotes? endnotes, Dictionary<int, int> numberMap)
    {
        var entries = new List<NoteEntry>();
        if (endnotes == null) return entries;

        foreach (var endnote in endnotes.Elements<Endnote>())
        {
            var idAttr = endnote.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
            if (string.IsNullOrEmpty(idAttr.Value) || !int.TryParse(idAttr.Value, out var id) || id <= 0) continue;

            var number = numberMap.TryGetValue(id, out var n) ? n : id;
            var entry = new NoteEntry
            {
                Id = id,
                Number = number,
                IsFootnote = false,
                ReferenceMark = "¹"
            };

            // 提取内容
            ExtractNoteContent(endnote, entry);
            entries.Add(entry);
        }

        return entries;
    }

    /// <summary>
    /// 提取注释内容
    /// </summary>
    private static void ExtractNoteContent(OpenXmlCompositeElement noteElement, NoteEntry entry)
    {
        var textBuilder = new System.Text.StringBuilder();
        
        // 查找 body 元素
        var body = noteElement.Elements().FirstOrDefault(e => 
            e.LocalName == "body" || e.LocalName == "Body");
        
        var contentElement = body ?? noteElement;
        
        foreach (var text in contentElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
        {
            textBuilder.Append(text.Text);
        }
        
        entry.Content = textBuilder.ToString().Trim();
    }

    /// <summary>
    /// 在 PDF 中渲染脚注（页面底部）
    /// </summary>
    public static void RenderFootnote(
        PdfContentByte canvas,
        NoteEntry note,
        float leftMargin,
        float rightMargin,
        float bottomMargin,
        float pageWidth,
        float availableHeight)
    {
        // 绘制分隔线
        var separatorY = bottomMargin + availableHeight;
        var separatorWidth = pageWidth * 0.3f;
        
        canvas.SaveState();
        canvas.SetLineWidth(0.5f);
        canvas.SetColorStroke(BaseColor.Gray);
        canvas.MoveTo(leftMargin, separatorY);
        canvas.LineTo(leftMargin + separatorWidth, separatorY);
        canvas.Stroke();
        canvas.RestoreState();

        // 渲染注释内容
        var noteFont = FontFactory.GetFont("STSong-Light", 9);
        var noteText = $"{note.Number}. {note.Content}";
        
        var noteParagraph = new iTextParagraph(noteText, noteFont);
        noteParagraph.SpacingBefore = 2f;
        noteParagraph.SpacingAfter = 4f;
        noteParagraph.IndentationLeft = 15f;
        
        // 使用 ColumnText 渲染到指定区域
        var column = new ColumnText(canvas);
        column.SetSimpleColumn(
            leftMargin,
            bottomMargin,
            rightMargin,
            separatorY - 2f
        );
        column.AddElement(noteParagraph);
        column.Go();
    }

    /// <summary>
    /// 在 PDF 末尾渲染尾注
    /// </summary>
    public static void RenderEndnotes(
        List<NoteEntry> endnotes,
        PdfDocument pdfDocument)
    {
        if (endnotes.Count == 0) return;

        // 添加新页面
        pdfDocument.NewPage();

        // 尾注标题
        var titleFont = FontFactory.GetFont("STSong-Light", 18, iTextFont.BOLD);
        var title = new iTextParagraph("尾注", titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 20f
        };
        pdfDocument.Add(title);

        // 尾注内容
        var contentFont = FontFactory.GetFont("STSong-Light", 10);
        
        foreach (var note in endnotes)
        {
            var notePara = new iTextParagraph();
            notePara.Font = contentFont;
            notePara.SpacingAfter = 8f;

            // 编号（上标）
            var numberChunk = new iTextChunk($"{note.Number}.", contentFont);
            numberChunk.SetTextRise(4f);
            notePara.Add(numberChunk);

            // 内容
            var contentChunk = new iTextChunk($" {note.Content}", contentFont);
            notePara.Add(contentChunk);

            pdfDocument.Add(notePara);
        }
    }

    /// <summary>
    /// 创建脚注引用标记（上标数字）
    /// </summary>
    public static Chunk CreateFootnoteReference(int number, float baseFontSize)
    {
        var font = FontFactory.GetFont("STSong-Light", baseFontSize * 0.7f);
        var chunk = new iTextChunk(number.ToString(), font);
        chunk.SetTextRise(baseFontSize * 0.35f);
        return chunk;
    }

    /// <summary>
    /// 自定义脚注编号格式
    /// </summary>
    public static string FormatNoteNumber(int number, string format = "arabic")
    {
        return format.ToLower() switch
        {
            "roman" => ToRomanNumerals(number),
            "alpha" => ToAlpha(number),
            "chinese" => ToChineseNumber(number),
            _ => number.ToString() // arabic
        };
    }

    /// <summary>
    /// 转换为罗马数字
    /// </summary>
    private static string ToRomanNumerals(int number)
    {
        if (number <= 0 || number > 3999) return number.ToString();
        
        var values = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var symbols = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        
        var result = new System.Text.StringBuilder();
        for (var i = 0; i < values.Length; i++)
        {
            while (number >= values[i])
            {
                number -= values[i];
                result.Append(symbols[i]);
            }
        }
        return result.ToString();
    }

    /// <summary>
    /// 转换为字母编号（A, B, C...）
    /// </summary>
    private static string ToAlpha(int number)
    {
        if (number <= 0) return number.ToString();
        
        var letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        var result = new System.Text.StringBuilder();
        
        while (number > 0)
        {
            number--;
            result.Insert(0, letters[number % 26]);
            number /= 26;
        }
        
        return result.ToString();
    }

    /// <summary>
    /// 转换为中文数字（一、二、三...）
    /// </summary>
    private static string ToChineseNumber(int number)
    {
        if (number <= 0) return number.ToString();
        
        var digits = "零一二三四五六七八九";
        var units = new[] { "", "十", "百", "千" };
        
        var numStr = number.ToString();
        var result = new System.Text.StringBuilder();
        var unitPos = 0;
        var needZero = false;
        
        for (var i = numStr.Length - 1; i >= 0; i--)
        {
            var digit = numStr[i] - '0';
            if (digit == 0)
            {
                needZero = true;
            }
            else
            {
                if (needZero)
                {
                    result.Insert(0, "零");
                    needZero = false;
                }
                result.Insert(0, units[unitPos]);
                result.Insert(0, digits[digit]);
            }
            unitPos++;
        }
        
        return result.ToString();
    }
}
