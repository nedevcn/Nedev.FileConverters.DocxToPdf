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
    /// Split string into directional runs and reverse those identified as RTL.
    /// This is a very simple approximation of the Unicode bidi algorithm.
    /// </summary>
    private static string ApplyBasicBidi(string text, string? language)
    {
        if (string.IsNullOrEmpty(text)) return text;
        bool IsRtlChar(char c)
        {
            var cat = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            return cat == System.Globalization.UnicodeCategory.OtherLetter &&
                   (c >= '\u0590' && c <= '\u08FF'); // Hebrew, Arabic ranges
        }
        var result = new System.Text.StringBuilder();
        int i = 0;
        while (i < text.Length)
        {
            bool rtl = IsRtlChar(text[i]);
            int j = i + 1;
            while (j < text.Length && IsRtlChar(text[j]) == rtl) j++;
            var segment = text.Substring(i, j - i);
            if (rtl)
            {
                char[] arr = segment.ToCharArray();
                Array.Reverse(arr);
                result.Append(arr);
            }
            else
            {
                result.Append(segment);
            }
            i = j;
        }
        return result.ToString();
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

                // resolve bidi if language is RTL
                var textContent = textNode.Text;
                // run-specific bidi handling: split into runs of same direction and reverse RTL runs
                textContent = ApplyBasicBidi(textContent, runPr?.Language?.Val);

                // embedded field resolution: match typical field pattern of "WORD ..." and call converter
                if (System.Text.RegularExpressions.Regex.IsMatch(textContent, @"^[A-Z]{2,}\b"))
                {
                    try
                    {
                        var resolved = new DocxToPdfConverter().ResolveField(textContent, _document, "");
                        if (resolved != null)
                            textContent = resolved;
                    }
                    catch { }
                }

                var fontObj = _fontHelper.GetFont(runPr);
                var chunk = new iTextChunk(textContent, fontObj);
                // detect anything mentioning vertical to override direction
                if (runPr != null && runPr.AnyAttrs)
                {
                    foreach (var attr in runPr.GetAttributes())
                    {
                        if (attr.LocalName.Contains("vert", StringComparison.OrdinalIgnoreCase))
                        {
                            chunk.DirectionOverride = TextDirection.Vertical;
                            break;
                        }
                    }
                }

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

                // gradient fill: average all defined stops, optionally weighting by position
                if (runPr?.GradientFill != null)
                {
                    var stops = runPr.GradientFill.Descendants<DocumentFormat.OpenXml.Drawing.GradientStop>().ToList();
                    if (stops.Count > 0)
                    {
                        long totalWeight = 0;
                        long sumR = 0, sumG = 0, sumB = 0;
                        foreach (var st in stops)
                        {
                            string? clr = st.Descendants<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
                            if (clr == null) continue;
                            int v = int.Parse(clr, System.Globalization.NumberStyles.HexNumber);
                            int r = (v >> 16) & 0xFF;
                            int g = (v >> 8) & 0xFF;
                            int b = v & 0xFF;
                            long pos = st.Position?.Value ?? 0;
                            long weight = pos > 0 ? pos : 1;
                            sumR += r * weight;
                            sumG += g * weight;
                            sumB += b * weight;
                            totalWeight += weight;
                        }
                        if (totalWeight > 0)
                        {
                            int r = (int)(sumR / totalWeight);
                            int g = (int)(sumG / totalWeight);
                            int b = (int)(sumB / totalWeight);
                            chunk.Font.Color = new BaseColor(r, g, b);
                        }
                    }
                }

                // character spacing if defined
                var cs = runPr?.GetFirstChild<DocumentFormat.OpenXml.Drawing.CharacterSpacing>()?.Val?.Value;
                if (cs != null && float.TryParse(cs, out var csVal))
                {
                    // characterSpacing is in thousandths of a point per OpenXML spec
                    chunk.CharSpacing = csVal / 1000f;
                }

                pdfPara.Add(chunk);
    }
}
