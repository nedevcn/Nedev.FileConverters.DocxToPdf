using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using Nedev.FileConverters.DocxToPdf.Rendering;
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
    private readonly DrawingMLRenderer _renderer;

    public DrawingMLConverter(WordprocessingDocument document, FontHelper fontHelper, ConvertOptions? options = null)
    {
        _document = document;
        _fontHelper = fontHelper;
        _renderer = new DrawingMLRenderer(document, options ?? ConvertOptions.Default);
    }

    /// <summary>
    /// 尝试将 Drawing 元素转换为 PDF 元素
    /// </summary>
    public IElement? ConvertDrawing(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, float pageWidth)
    {
        try
        {
            // if the drawing contains text but no embedded picture we prefer returning
            // a paragraph rather than rasterizing the whole graphic. this covers
            // textboxes and many SmartArt nodes where the visual content is
            // purely textual.
            bool hasText = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                                .Any(t => !string.IsNullOrWhiteSpace(t.Text));
            bool hasBlip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Any();
            if (hasText && !hasBlip)
            {
                var textElem = ExtractTextFromDrawing(drawing, pageWidth);
                if (textElem != null)
                    return textElem;
            }

            // 首先尝试使用 SkiaSharp 渲染器渲染为图片
            var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
            if (extent != null)
            {
                var widthPx = (int)EMU.ToPixels(extent.Cx?.Value ?? 914400);
                var heightPx = (int)EMU.ToPixels(extent.Cy?.Value ?? 914400);

                // 限制最大尺寸
                if (widthPx > 2000) widthPx = 2000;
                if (heightPx > 2000) heightPx = 2000;

                var pngBytes = _renderer.RenderToPng(drawing, widthPx, heightPx);
                if (pngBytes != null && pngBytes.Length > 0)
                {
                    var pdfImage = iTextImage.GetInstance(pngBytes);

                    // 设置尺寸
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

                    // 检查是否是浮动对象
                    var anchor = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
                    if (anchor != null)
                    {
                        return new FloatingObject(pdfImage);
                    }

                    return pdfImage;
                }
            }

            // 如果渲染失败，回退到文本提取
            return ExtractTextFromDrawing(drawing, pageWidth);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[DrawingMLConverter] Failed to convert drawing: {ex.Message}");
            return ExtractTextFromDrawing(drawing, pageWidth);
        }
    }

    /// <summary>
    /// 从 Drawing 提取文本
    /// </summary>
    private IElement? ExtractTextFromDrawing(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, float pageWidth)
    {
        var graphicData = drawing.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().FirstOrDefault();
        if (graphicData == null) return null;

        // look for paragraphs inside the graphic; must have at least one run of text
        var paragraphs = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();
        if (paragraphs.Count == 0) return null;

        var pdfPara = new iTextParagraph();

        foreach (var para in paragraphs)
        {
            // determine alignment if specified
            var align = para.ParagraphProperties?.Alignment?.Value;
            if (align != null)
            {
                switch (align.Value)
                {
                    case DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center:
                        pdfPara.Alignment = Element.ALIGN_CENTER;
                        break;
                    case DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right:
                        pdfPara.Alignment = Element.ALIGN_RIGHT;
                        break;
                    case DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Justified:
                        pdfPara.Alignment = Element.ALIGN_JUSTIFIED;
                        break;
                    default:
                        pdfPara.Alignment = Element.ALIGN_LEFT;
                        break;
                }
            }

            foreach (var run in para.Descendants<DocumentFormat.OpenXml.Drawing.Run>())
            {
                var textNode = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                if (textNode == null || string.IsNullOrWhiteSpace(textNode.Text)) continue;

                var runPr = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
                var fontObj = _fontHelper.GetFont(runPr);
                var chunk = new iTextChunk(textNode.Text, fontObj);

                // underline/strike
                if (runPr?.Underline != null && runPr.Underline.Val != DocumentFormat.OpenXml.Drawing.TextUnderlineValues.None)
                {
                    chunk.SetUnderline(0.1f, -1f);
                }
                if (runPr?.Strike != null)
                {
                    chunk.Font.Style |= iTextFont.STRIKETHRU;
                }

                // small caps: render uppercase while shrinking size
                if (runPr?.SmallCaps != null)
                {
                    chunk.Content = chunk.Content.ToUpperInvariant();
                    chunk.Font.Size *= 0.8f;
                }

                // color gradient: choose first stop if gradient fill present
                if (runPr?.GradientFill != null)
                {
                    var stop = runPr.GradientFill.Descendants<DocumentFormat.OpenXml.Drawing.GradientStop>().FirstOrDefault();
                    var clr = stop?.Descendants<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
                    if (!string.IsNullOrEmpty(clr))
                    {
                        try { chunk.Font.Color = new BaseColor(int.Parse(clr, System.Globalization.NumberStyles.HexNumber)); } catch { }
                    }
                }

                // TODO: character spacing (kerning/tracking) could be applied here when
                // DrawingML defines it; would require a new property on Chunk or Font.

                pdfPara.Add(chunk);
    }
}
