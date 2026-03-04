using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.DocxToPdf.PdfEngine;
using iTextFont = Nedev.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.DocxToPdf.PdfEngine.Paragraph;
using iTextChunk = Nedev.DocxToPdf.PdfEngine.Chunk;

namespace Nedev.DocxToPdf.Helpers;

/// <summary>
/// 修订标记（Track Changes）处理器
/// </summary>
public static class RevisionHandler
{
    /// <summary>
    /// 修订类型
    /// </summary>
    public enum RevisionType
    {
        None,
        Insertion,
        Deletion,
        Formatting,
        Comment
    }

    /// <summary>
    /// 修订信息
    /// </summary>
    public class RevisionInfo
    {
        public RevisionType Type { get; set; } = RevisionType.None;
        public string? Author { get; set; }
        public System.DateTime? Date { get; set; }
        public string? CommentText { get; set; }
    }

    /// <summary>
    /// 检查 Run 是否属于插入的修订
    /// </summary>
    public static bool IsInserted(Run run)
    {
        // 检查 RunProperties 中的插入标记
        var runProps = run.RunProperties;
        if (runProps != null)
        {
            // 检查 insert 属性
            foreach (var attr in runProps.GetAttributes())
            {
                if ((attr.LocalName == "ins" || attr.LocalName == "inserted") &&
                    (attr.Value == "1" || attr.Value == "true"))
                {
                    return true;
                }
            }
        }

        // 检查祖先元素中的插入标记
        var parentInsert = run.Ancestors<Inserted>().FirstOrDefault();
        return parentInsert != null;
    }

    /// <summary>
    /// 检查 Run 是否属于删除的修订
    /// </summary>
    public static bool IsDeleted(Run run)
    {
        // 检查 RunProperties 中的删除标记
        var runProps = run.RunProperties;
        if (runProps != null)
        {
            foreach (var attr in runProps.GetAttributes())
            {
                if ((attr.LocalName == "del" || attr.LocalName == "deleted") &&
                    (attr.Value == "1" || attr.Value == "true"))
                {
                    return true;
                }
            }
        }

        // 检查祖先元素中的删除标记
        var parentDelete = run.Ancestors<Deleted>().FirstOrDefault();
        return parentDelete != null;
    }

    /// <summary>
    /// 检查段落是否包含修订
    /// </summary>
    public static bool HasRevisions(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
    {
        return paragraph.Descendants<Inserted>().Any() ||
               paragraph.Descendants<Deleted>().Any();
    }

    /// <summary>
    /// 获取修订信息
    /// </summary>
    public static RevisionInfo GetRevisionInfo(Run run)
    {
        var info = new RevisionInfo();

        // 检查插入
        if (IsInserted(run))
        {
            info.Type = RevisionType.Insertion;
            var insert = run.Ancestors<Inserted>().FirstOrDefault();
            if (insert != null)
            {
                info.Author = insert.Author?.Value;
                if (insert.Date?.HasValue == true)
                {
                    info.Date = insert.Date.Value;
                }
            }
            return info;
        }

        // 检查删除
        if (IsDeleted(run))
        {
            info.Type = RevisionType.Deletion;
            var delete = run.Ancestors<Deleted>().FirstOrDefault();
            if (delete != null)
            {
                info.Author = delete.Author?.Value;
                if (delete.Date?.HasValue == true)
                {
                    info.Date = delete.Date.Value;
                }
            }
            return info;
        }

        return info;
    }

    /// <summary>
    /// 应用修订样式到 Chunk
    /// </summary>
    public static void ApplyRevisionStyle(Chunk chunk, RevisionInfo info)
    {
        if (info.Type == RevisionType.None) return;

        var font = chunk.Font;
        if (font == null) return;

        switch (info.Type)
        {
            case RevisionType.Insertion:
                // 插入内容：绿色
                var insertFont = new iTextFont(font.Family, font.Size, font.Style, new BaseColor(0, 128, 0));
                chunk.Font = insertFont;
                break;

            case RevisionType.Deletion:
                // 删除内容：红色删除线
                var deleteFont = new iTextFont(font.Family, font.Size, font.Style | iTextFont.STRIKETHRU, new BaseColor(200, 0, 0));
                chunk.Font = deleteFont;
                break;
        }
    }

    /// <summary>
    /// 收集文档中的所有修订
    /// </summary>
    public static List<(RevisionType Type, string Content, string? Author, System.DateTime? Date)> CollectRevisions(Body body)
    {
        var revisions = new List<(RevisionType, string, string?, System.DateTime?)>();

        foreach (var run in body.Descendants<Run>())
        {
            var info = GetRevisionInfo(run);
            if (info.Type != RevisionType.None)
            {
                var content = string.Join("", run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                if (!string.IsNullOrEmpty(content))
                {
                    revisions.Add((info.Type, content, info.Author, info.Date));
                }
            }
        }

        return revisions;
    }

    /// <summary>
    /// 在 PDF 中添加修订汇总页
    /// </summary>
    public static void AddRevisionsSummaryPage(Body body, PdfDocument pdfDocument)
    {
        var revisions = CollectRevisions(body);
        if (revisions.Count == 0) return;

        // 添加新页面
        pdfDocument.NewPage();

        // 添加标题
        var titleFont = FontFactory.GetFont("STSong-Light", 18, iTextFont.BOLD);
        var title = new iTextParagraph("修订记录", titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 20f
        };
        pdfDocument.Add(title);

        // 列出所有修订
        var contentFont = FontFactory.GetFont("STSong-Light", 11);

        foreach (var (type, content, author, date) in revisions)
        {
            var typeText = type == RevisionType.Insertion ? "[插入]" : "[删除]";
            var typeColor = type == RevisionType.Insertion ? new BaseColor(0, 128, 0) : new BaseColor(200, 0, 0);

            var revisionPara = new iTextParagraph();
            revisionPara.Font = contentFont;

            // 类型标记
            var typeChunk = new iTextChunk(typeText, new iTextFont(contentFont.Family, 11, iTextFont.BOLD, typeColor));
            revisionPara.Add(typeChunk);

            // 作者和时间
            if (!string.IsNullOrEmpty(author))
            {
                revisionPara.Add(new iTextChunk($" - 作者：{author}", contentFont));
            }
            if (date.HasValue)
            {
                revisionPara.Add(new iTextChunk($" ({date.Value:yyyy-MM-dd HH:mm})", contentFont));
            }

            // 内容
            revisionPara.Add(new iTextChunk($": {content}\n\n", contentFont));
            revisionPara.SpacingAfter = 8f;

            pdfDocument.Add(revisionPara);
        }
    }
}
