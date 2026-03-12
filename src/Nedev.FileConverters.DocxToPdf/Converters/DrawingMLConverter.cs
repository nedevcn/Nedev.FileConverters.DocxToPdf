using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;
using iTextImage = Nedev.FileConverters.DocxToPdf.PdfEngine.Image;

namespace Nedev.FileConverters.DocxToPdf.Converters;

/// <summary>
/// DrawingML 转换器 - 将 DrawingML 形状、文本框和图片转换为 PDF 元素
/// </summary>
public class DrawingMLConverter
{
    private readonly WordprocessingDocument _document;
    private readonly FontHelper _fontHelper;

    public DrawingMLConverter(WordprocessingDocument document, FontHelper fontHelper)
    {
        _document = document;
        _fontHelper = fontHelper;
    }

    /// <summary>
    /// 尝试将 Drawing 元素转换为 PDF 元素
    /// </summary>
    public IElement? ConvertDrawing(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, float pageWidth)
    {
        try
        {
            // 查找 GraphicData
            var graphicData = drawing.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().FirstOrDefault();
            if (graphicData == null) return null;

            // 检查是否是图片 (通过查找 Blip)
            var blip = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            if (blip != null)
            {
                return ConvertBlip(blip, drawing, pageWidth);
            }

            // 检查是否是形状 (通过查找 Shape)
            var shape = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.Shape>().FirstOrDefault();
            if (shape != null)
            {
                return ConvertShape(shape, pageWidth);
            }

            // 提取所有文本作为回退
            return ExtractTextFromGraphicData(graphicData, pageWidth);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[DrawingMLConverter] Failed to convert drawing: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 转换图片 (Blip)
    /// </summary>
    private IElement? ConvertBlip(DocumentFormat.OpenXml.Drawing.Blip blip, DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, float pageWidth)
    {
        var embedId = blip.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return null;

        try
        {
            var imagePart = _document.MainDocumentPart?.GetPartById(embedId) as ImagePart;
            if (imagePart == null) return null;

            using var stream = imagePart.GetStream();
            using var memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            var imageBytes = memoryStream.ToArray();

            if (imageBytes.Length == 0) return null;

            var pdfImage = iTextImage.GetInstance(imageBytes);

            // 获取尺寸
            var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
            if (extent != null)
            {
                var widthPt = StyleHelper.EmuToPoints(extent.Cx?.Value ?? 0);
                var heightPt = StyleHelper.EmuToPoints(extent.Cy?.Value ?? 0);

                if (widthPt > 0 && heightPt > 0)
                {
                    // 限制最大宽度
                    if (widthPt > pageWidth)
                    {
                        var ratio = pageWidth / widthPt;
                        widthPt = pageWidth;
                        heightPt *= ratio;
                    }
                    pdfImage.ScaleAbsolute(widthPt, heightPt);
                }
            }

            // 检查是否是浮动对象
            var anchor = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
            if (anchor != null)
            {
                return new FloatingObject(pdfImage);
            }

            return pdfImage;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[DrawingMLConverter] Failed to convert blip: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// 转换单个形状
    /// </summary>
    private IElement? ConvertShape(DocumentFormat.OpenXml.Drawing.Shape shape, float pageWidth)
    {
        // 获取形状中的文本
        var texts = new List<string>();

        // 查找 TextBody
        var txBody = shape.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>().FirstOrDefault();
        if (txBody != null)
        {
            foreach (var para in txBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                var paraText = string.Join("", para.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
                if (!string.IsNullOrWhiteSpace(paraText))
                {
                    texts.Add(paraText);
                }
            }
        }

        if (texts.Count == 0) return null;

        // 创建 PDF 段落
        var pdfPara = new iTextParagraph();
        var font = _fontHelper.GetFont(12f);

        foreach (var text in texts)
        {
            pdfPara.Add(new iTextChunk(text, font));
        }

        return pdfPara;
    }

    /// <summary>
    /// 从 GraphicData 提取文本
    /// </summary>
    private IElement? ExtractTextFromGraphicData(DocumentFormat.OpenXml.Drawing.GraphicData graphicData, float pageWidth)
    {
        var texts = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
            .Select(t => t.Text)
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .ToList();

        if (texts.Count == 0) return null;

        var pdfPara = new iTextParagraph();
        var font = _fontHelper.GetFont(12f);

        foreach (var text in texts)
        {
            pdfPara.Add(new iTextChunk(text, font));
            pdfPara.Add(new iTextChunk(" ", font));
        }

        return pdfPara;
    }
}
