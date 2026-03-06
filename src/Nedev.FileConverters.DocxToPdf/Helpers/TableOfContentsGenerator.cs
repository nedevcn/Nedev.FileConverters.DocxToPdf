using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

/// <summary>
/// ?????
/// </summary>
public static class TableOfContentsGenerator
{
    /// <summary>
    /// ????
    /// </summary>
    public class TOCEntry
    {
        public string Title { get; set; } = "";
        public int Level { get; set; }
        public int PageNumber { get; set; }
        public string? BookmarkId { get; set; }
    }

    /// <summary>
    /// ??????????
    /// </summary>
    public static List<TOCEntry> ExtractTOC(Body body)
    {
        var entries = new List<TOCEntry>();
        var headingStyles = new[] { "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Heading7", "Heading8", "Heading9" };
        
        foreach (var paragraph in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
        {
            var paraProps = paragraph.ParagraphProperties;
            var styleId = paraProps?.ParagraphStyleId?.Val?.Value;
            
            if (string.IsNullOrEmpty(styleId)) continue;
            
            // ?????????
            var headingLevel = GetHeadingLevel(styleId);
            if (headingLevel.HasValue)
            {
                var title = string.Join("", paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text)).Trim();
                if (!string.IsNullOrEmpty(title))
                {
                    // ????(???)
                    var bookmarkStart = paragraph.Descendants<BookmarkStart>().FirstOrDefault();
                    var bookmarkId = bookmarkStart?.Name?.Value;
                    
                    entries.Add(new TOCEntry
                    {
                        Title = title,
                        Level = headingLevel.Value,
                        BookmarkId = bookmarkId
                    });
                }
            }
        }
        
        return entries;
    }

    /// <summary>
    /// ??????
    /// </summary>
    private static int? GetHeadingLevel(string styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId)) return null;
        
        var lower = styleId.ToLowerInvariant();
        if (lower.StartsWith("heading"))
        {
            var numPart = lower.Substring("heading".Length);
            if (int.TryParse(numPart, out var level) && level >= 1 && level <= 9)
                return level;
        }
        
        // ????????
        if (lower.Contains("??"))
        {
            if (lower.Contains("1") || lower.Contains("?")) return 1;
            if (lower.Contains("2") || lower.Contains("?")) return 2;
            if (lower.Contains("3") || lower.Contains("?")) return 3;
        }
        
        return null;
    }

    /// <summary>
    /// ? PDF ??????
    /// </summary>
    public static void GenerateTOC(
        List<TOCEntry> entries,
        PdfDocument pdfDocument,
        PdfWriter writer,
        int startPage = 1)
    {
        if (entries.Count == 0) return;

        // ?????
        pdfDocument.NewPage();

        // ????
        var titleFont = FontFactory.GetFont("STSong-Light", 24, iTextFont.BOLD);
        var title = new iTextParagraph("??", titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 30f
        };
        pdfDocument.Add(title);

        // ??????
        var contentFont = FontFactory.GetFont("STSong-Light", 12);
        var leaderFont = FontFactory.GetFont("STSong-Light", 10);

        // ??????
        foreach (var entry in entries)
        {
            // ????
            var indent = (entry.Level - 1) * 20f;
            
            // ????
            var para = new iTextParagraph();
            para.IndentationLeft = indent;
            para.SpacingAfter = 4f;

            // ????
            var titleChunk = new iTextChunk(entry.Title, contentFont);
            if (entry.BookmarkId != null)
            {
                // ??????
                var destination = new PdfDestination(PdfDestination.FIT);
                // ??:????????????????
                titleChunk.SetAnchor(entry.BookmarkId);
            }
            para.Add(titleChunk);

            // ???(??)- ????,?????
            para.Add(new iTextChunk(" ... ", contentFont));

            // ??(???????,??????????????)
            var pageChunk = new iTextChunk(" ?", contentFont);
            para.Add(pageChunk);

            pdfDocument.Add(para);
        }
    }

    /// <summary>
    /// ??????(????????)
    /// </summary>
    public static void UpdateTOCPageNumbers(
        PdfDocument pdfDocument,
        PdfWriter writer,
        List<TOCEntry> entries,
        int tocStartPage)
    {
        // ????????????????,??????????
        // ??????????
        // ???????,???????
    }

    /// <summary>
    /// ?? PDF ????
    /// </summary>
    public static void CreateBookmarks(
        List<TOCEntry> entries,
        PdfWriter writer,
        Dictionary<string, int> bookmarkPageMap)
    {
        var rootOutline = writer.RootOutline;
        var outlineByLevel = new Dictionary<int, PdfOutline>();

        foreach (var entry in entries)
        {
            var level = Math.Clamp(entry.Level, 1, 6);
            
            // ????
            var pageNumber = bookmarkPageMap.TryGetValue(entry.BookmarkId ?? entry.Title, out var pg) ? pg : 1;
            
            // ????
            var dest = new PdfDestination(PdfDestination.FIT);
            // iTextSharp.LGPLv2.Core ?????????,????????

            // ??????
            PdfOutline parent;
            if (level == 1)
            {
                parent = rootOutline;
            }
            else
            {
                outlineByLevel.TryGetValue(level - 1, out parent);
                parent ??= rootOutline;
            }

            // ??????
            var outline = new PdfOutline(parent, dest, entry.Title, level <= 2);
            outlineByLevel[level] = outline;
        }
    }
}

/// <summary>
/// ?????(???????)
/// </summary>
public class DottedLineSeparator
{
    public float Leading { get; set; } = 0;
    
    public void Draw(PdfContentByte canvas, float minX, float maxX, float y)
    {
        var dotSpacing = 4f;
        var dotSize = 1f;
        
        canvas.SaveState();
        canvas.SetColorFill(BaseColor.Gray);
        
        var x = minX;
        while (x < maxX)
        {
            canvas.Rectangle(x, y - dotSize / 2, dotSize, dotSize);
            x += dotSpacing;
        }
        
        canvas.Fill();
        canvas.RestoreState();
    }
}
