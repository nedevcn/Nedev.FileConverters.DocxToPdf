using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

/// <summary>
/// Word ????? - ? Word ????? PDF ??
/// </summary>
public static class CommentExporter
{
    /// <summary>
    /// ????????????
    /// </summary>
    public static void AddCommentsSummaryPage(WordprocessingDocument docxDocument, PdfDocument pdfDocument)
    {
        var commentsPart = docxDocument.MainDocumentPart?.GetPartsOfType<WordprocessingCommentsPart>().FirstOrDefault();
        if (commentsPart == null) return;

        var comments = commentsPart.Comments?.Elements<Comment>();
        if (comments == null || !comments.Any()) return;

        // ?????
        pdfDocument.NewPage();

        // ????
        var titleFont = FontFactory.GetFont("STSong-Light", 18, iTextFont.BOLD);
        var title = new iTextParagraph("????", titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 20f
        };
        pdfDocument.Add(title);

        // ??????
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
            commentPara.Add(new Chunk($"??:{author}", commentFont));
            
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
