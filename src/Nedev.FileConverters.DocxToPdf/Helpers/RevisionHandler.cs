using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

/// <summary>
/// ????(Track Changes)???
/// </summary>
public static class RevisionHandler
{
    /// <summary>
    /// ????
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
    /// ????
    /// </summary>
    public class RevisionInfo
    {
        public RevisionType Type { get; set; } = RevisionType.None;
        public string? Author { get; set; }
        public System.DateTime? Date { get; set; }
        public string? CommentText { get; set; }
    }

    /// <summary>
    /// ?? Run ?????????
    /// </summary>
    public static bool IsInserted(Run run)
    {
        // ?? RunProperties ??????
        var runProps = run.RunProperties;
        if (runProps != null)
        {
            // ?? insert ??
            foreach (var attr in runProps.GetAttributes())
            {
                if ((attr.LocalName == "ins" || attr.LocalName == "inserted") &&
                    (attr.Value == "1" || attr.Value == "true"))
                {
                    return true;
                }
            }
        }

        // ????????????
        var parentInsert = run.Ancestors<Inserted>().FirstOrDefault();
        return parentInsert != null;
    }

    /// <summary>
    /// ?? Run ?????????
    /// </summary>
    public static bool IsDeleted(Run run)
    {
        // ?? RunProperties ??????
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

        // ????????????
        var parentDelete = run.Ancestors<Deleted>().FirstOrDefault();
        return parentDelete != null;
    }

    /// <summary>
    /// ??????????
    /// </summary>
    public static bool HasRevisions(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
    {
        return paragraph.Descendants<Inserted>().Any() ||
               paragraph.Descendants<Deleted>().Any();
    }

    /// <summary>
    /// ??????
    /// </summary>
    public static RevisionInfo GetRevisionInfo(Run run)
    {
        var info = new RevisionInfo();

        // ????
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

        // ????
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
    /// ??????? Chunk
    /// </summary>
    public static void ApplyRevisionStyle(Chunk chunk, RevisionInfo info)
    {
        if (info.Type == RevisionType.None) return;

        var font = chunk.Font;
        if (font == null) return;

        switch (info.Type)
        {
            case RevisionType.Insertion:
                // ????:??
                var insertFont = new iTextFont(font.Family, font.Size, font.Style, new BaseColor(0, 128, 0));
                chunk.Font = insertFont;
                break;

            case RevisionType.Deletion:
                // ????:?????
                var deleteFont = new iTextFont(font.Family, font.Size, font.Style | iTextFont.STRIKETHRU, new BaseColor(200, 0, 0));
                chunk.Font = deleteFont;
                break;
        }
    }

    /// <summary>
    /// ??????????
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
    /// ? PDF ????????
    /// </summary>
    public static void AddRevisionsSummaryPage(Body body, PdfDocument pdfDocument)
    {
        var revisions = CollectRevisions(body);
        if (revisions.Count == 0) return;

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
        var contentFont = FontFactory.GetFont("STSong-Light", 11);

        foreach (var (type, content, author, date) in revisions)
        {
            var typeText = type == RevisionType.Insertion ? "[??]" : "[??]";
            var typeColor = type == RevisionType.Insertion ? new BaseColor(0, 128, 0) : new BaseColor(200, 0, 0);

            var revisionPara = new iTextParagraph();
            revisionPara.Font = contentFont;

            // ????
            var typeChunk = new iTextChunk(typeText, new iTextFont(contentFont.Family, 11, iTextFont.BOLD, typeColor));
            revisionPara.Add(typeChunk);

            // ?????
            if (!string.IsNullOrEmpty(author))
            {
                revisionPara.Add(new iTextChunk($" - ??:{author}", contentFont));
            }
            if (date.HasValue)
            {
                revisionPara.Add(new iTextChunk($" ({date.Value:yyyy-MM-dd HH:mm})", contentFont));
            }

            // ??
            revisionPara.Add(new iTextChunk($": {content}\n\n", contentFont));
            revisionPara.SpacingAfter = 8f;

            pdfDocument.Add(revisionPara);
        }
    }
}
