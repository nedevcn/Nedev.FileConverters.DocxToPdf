using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

/// <summary>
/// ??????????
/// </summary>
public static class FootnoteEndnoteEnhancer
{
    /// <summary>
    /// ??/????
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
    /// ???????????
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

            // ????
            ExtractNoteContent(footnote, entry);
            entries.Add(entry);
        }

        return entries;
    }

    /// <summary>
    /// ???????????
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
                ReferenceMark = "�"
            };

            // ????
            ExtractNoteContent(endnote, entry);
            entries.Add(entry);
        }

        return entries;
    }

    /// <summary>
    /// ??????
    /// </summary>
    private static void ExtractNoteContent(OpenXmlCompositeElement noteElement, NoteEntry entry)
    {
        var textBuilder = new System.Text.StringBuilder();
        
        // ?? body ??
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
    /// ? PDF ?????(????)
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
        // ?????
        var separatorY = bottomMargin + availableHeight;
        var separatorWidth = pageWidth * 0.3f;
        
        canvas.SaveState();
        canvas.SetLineWidth(0.5f);
        canvas.SetColorStroke(BaseColor.Gray);
        canvas.MoveTo(leftMargin, separatorY);
        canvas.LineTo(leftMargin + separatorWidth, separatorY);
        canvas.Stroke();
        canvas.RestoreState();

        // ??????
        var noteFont = FontFactory.GetFont("STSong-Light", 9);
        var noteText = $"{note.Number}. {note.Content}";
        
        var noteParagraph = new iTextParagraph(noteText, noteFont);
        noteParagraph.SpacingBefore = 2f;
        noteParagraph.SpacingAfter = 4f;
        noteParagraph.IndentationLeft = 15f;
        
        // ?? ColumnText ???????
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
    /// ? PDF ??????
    /// </summary>
    public static void RenderEndnotes(
        List<NoteEntry> endnotes,
        PdfDocument pdfDocument)
    {
        if (endnotes.Count == 0) return;

        // ?????
        pdfDocument.NewPage();

        // ????
        var titleFont = FontFactory.GetFont("STSong-Light", 18, iTextFont.BOLD);
        var title = new iTextParagraph("??", titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 20f
        };
        pdfDocument.Add(title);

        // ????
        var contentFont = FontFactory.GetFont("STSong-Light", 10);
        
        foreach (var note in endnotes)
        {
            var notePara = new iTextParagraph();
            notePara.Font = contentFont;
            notePara.SpacingAfter = 8f;

            // ??(??)
            var numberChunk = new iTextChunk($"{note.Number}.", contentFont);
            numberChunk.SetTextRise(4f);
            notePara.Add(numberChunk);

            // ??
            var contentChunk = new iTextChunk($" {note.Content}", contentFont);
            notePara.Add(contentChunk);

            pdfDocument.Add(notePara);
        }
    }

    /// <summary>
    /// ????????(????)
    /// </summary>
    public static Chunk CreateFootnoteReference(int number, float baseFontSize)
    {
        var font = FontFactory.GetFont("STSong-Light", baseFontSize * 0.7f);
        var chunk = new iTextChunk(number.ToString(), font);
        chunk.SetTextRise(baseFontSize * 0.35f);
        return chunk;
    }

    /// <summary>
    /// ?????????
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
    /// ???????
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
    /// ???????(A, B, C...)
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
    /// ???????(?????...)
    /// </summary>
    private static string ToChineseNumber(int number)
    {
        if (number <= 0) return number.ToString();
        
        var digits = "??????????";
        var units = new[] { "", "?", "?", "?" };
        
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
                    result.Insert(0, "?");
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
