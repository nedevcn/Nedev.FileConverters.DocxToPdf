using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.DocxToPdf.PdfEngine;
using iTextFont = Nedev.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.DocxToPdf.PdfEngine.Paragraph;
using iTextChunk = Nedev.DocxToPdf.PdfEngine.Chunk;

namespace Nedev.DocxToPdf.Helpers;

/// <summary>
/// 目录生成器
/// </summary>
public static class TableOfContentsGenerator
{
    /// <summary>
    /// 目录条目
    /// </summary>
    public class TOCEntry
    {
        public string Title { get; set; } = "";
        public int Level { get; set; }
        public int PageNumber { get; set; }
        public string? BookmarkId { get; set; }
    }

    /// <summary>
    /// 从文档中提取目录结构
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
            
            // 检查是否为标题样式
            var headingLevel = GetHeadingLevel(styleId);
            if (headingLevel.HasValue)
            {
                var title = string.Join("", paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text)).Trim();
                if (!string.IsNullOrEmpty(title))
                {
                    // 查找书签（如果有）
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
    /// 获取标题级别
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
        
        // 检查中文标题样式
        if (lower.Contains("标题"))
        {
            if (lower.Contains("1") || lower.Contains("一")) return 1;
            if (lower.Contains("2") || lower.Contains("二")) return 2;
            if (lower.Contains("3") || lower.Contains("三")) return 3;
        }
        
        return null;
    }

    /// <summary>
    /// 在 PDF 中生成目录页
    /// </summary>
    public static void GenerateTOC(
        List<TOCEntry> entries,
        PdfDocument pdfDocument,
        PdfWriter writer,
        int startPage = 1)
    {
        if (entries.Count == 0) return;

        // 添加目录页
        pdfDocument.NewPage();

        // 目录标题
        var titleFont = FontFactory.GetFont("STSong-Light", 24, iTextFont.BOLD);
        var title = new iTextParagraph("目录", titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 30f
        };
        pdfDocument.Add(title);

        // 目录内容字体
        var contentFont = FontFactory.GetFont("STSong-Light", 12);
        var leaderFont = FontFactory.GetFont("STSong-Light", 10);

        // 生成目录条目
        foreach (var entry in entries)
        {
            // 计算缩进
            var indent = (entry.Level - 1) * 20f;
            
            // 创建段落
            var para = new iTextParagraph();
            para.IndentationLeft = indent;
            para.SpacingAfter = 4f;

            // 标题文本
            var titleChunk = new iTextChunk(entry.Title, contentFont);
            if (entry.BookmarkId != null)
            {
                // 添加内部链接
                var destination = new PdfDestination(PdfDestination.FIT);
                // 注意：这里需要在书签位置设置实际的页码
                titleChunk.SetAnchor(entry.BookmarkId);
            }
            para.Add(titleChunk);

            // 前导符（点线）- 简化处理，使用省略号
            para.Add(new iTextChunk(" ... ", contentFont));

            // 页码（暂时显示为问号，实际页码需要在布局完成后确定）
            var pageChunk = new iTextChunk(" ?", contentFont);
            para.Add(pageChunk);

            pdfDocument.Add(para);
        }
    }

    /// <summary>
    /// 更新目录页码（在文档关闭前调用）
    /// </summary>
    public static void UpdateTOCPageNumbers(
        PdfDocument pdfDocument,
        PdfWriter writer,
        List<TOCEntry> entries,
        int tocStartPage)
    {
        // 这个方法需要在文档布局完成后调用，用于更新目录中的页码
        // 需要在两遍处理中完成
        // 第一遍收集页码，第二遍更新目录
    }

    /// <summary>
    /// 创建 PDF 书签大纲
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
            
            // 获取页码
            var pageNumber = bookmarkPageMap.TryGetValue(entry.BookmarkId ?? entry.Title, out var pg) ? pg : 1;
            
            // 创建目标
            var dest = new PdfDestination(PdfDestination.FIT);
            // iTextSharp.LGPLv2.Core 不支持直接设置页码，需要使用间接引用

            // 找到父级大纲
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

            // 创建大纲条目
            var outline = new PdfOutline(parent, dest, entry.Title, level <= 2);
            outlineByLevel[level] = outline;
        }
    }
}

/// <summary>
/// 点线分隔符（用于目录前导符）
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
