using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.DocxToPdf.PdfEngine;
using iTextFont = Nedev.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.DocxToPdf.PdfEngine.Paragraph;

namespace Nedev.DocxToPdf.Helpers;

/// <summary>
/// Word 批注导出器 - 将 Word 批注转换为 PDF 注释
/// </summary>
public static class CommentExporter
{
    /// <summary>
    /// 在文档末尾添加批注汇总页
    /// </summary>
    public static void AddCommentsSummaryPage(WordprocessingDocument docxDocument, PdfDocument pdfDocument)
    {
        var commentsPart = docxDocument.MainDocumentPart?.GetPartsOfType<WordprocessingCommentsPart>().FirstOrDefault();
        if (commentsPart == null) return;

        var comments = commentsPart.Comments?.Elements<Comment>();
        if (comments == null || !comments.Any()) return;

        // 添加新页面
        pdfDocument.NewPage();

        // 添加标题
        var titleFont = FontFactory.GetFont("STSong-Light", 18, iTextFont.BOLD);
        var title = new iTextParagraph("批注汇总", titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 20f
        };
        pdfDocument.Add(title);

        // 列出所有批注
        var commentFont = FontFactory.GetFont("STSong-Light", 11);
        var index = 1;
        
        foreach (var comment in comments)
        {
            var commentText = ExtractCommentText(comment);
            var author = comment.Author?.Value ?? "Unknown";
            var date = comment.Date?.Value;

            var commentPara = new iTextParagraph();
            commentPara.Font = commentFont;
            commentPara.Add(new Chunk($"[{index}] ", commentFont));
            commentPara.Add(new Chunk($"作者：{author}", commentFont));
            
            if (date.HasValue)
            {
                commentPara.Add(new Chunk($" ({date.Value:yyyy-MM-dd HH:mm})", commentFont));
            }
            
            commentPara.Add(new Chunk($": {commentText}\n\n", commentFont));
            commentPara.SpacingAfter = 10f;
            
            pdfDocument.Add(commentPara);
            index++;
        }
    }

    private static string ExtractCommentText(Comment comment)
    {
        var textBuilder = new System.Text.StringBuilder();
        
        foreach (var element in comment.Descendants())
        {
            if (element is DocumentFormat.OpenXml.Wordprocessing.Text text)
            {
                textBuilder.Append(text.Text);
            }
        }
        
        return textBuilder.ToString().Trim();
    }
}
