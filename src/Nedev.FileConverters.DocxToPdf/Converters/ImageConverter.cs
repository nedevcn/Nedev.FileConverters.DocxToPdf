using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Helpers;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextImage = Nedev.FileConverters.DocxToPdf.PdfEngine.Image;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using System.Text;
using SkiaSharp;
using Nedev.FileConverters.DocxToPdf.Rasterization;

namespace Nedev.FileConverters.DocxToPdf.Converters;

/// <summary>
/// ????????
/// </summary>
public enum WrappingStyle
{
    Inline,         // ????
    InFrontOfText,  // ??????
    BehindText,     // ??????
    TopAndBottom,   // ???
    Square,         // ???
    Tight,          // ???(?? Square ??)
    Through         // ???(?? Square ??)
}

/// <summary>
/// ??????
/// </summary>
public class FloatingObject : IElement
{
    public iTextImage Image { get; set; }
    public WrappingStyle Wrapping { get; set; }
    public float Left { get; set; }
    public float Top { get; set; } // ??? Top ???????????,?????????????
    public bool PositionIsAbsolute { get; set; } // ??????????????
    public float Width => Image.ScaledWidth;
    public float Height => Image.ScaledHeight;
    
    // IElement ????
    public int Type => -100; // Custom type
    public bool IsContent() => true;
    public bool IsNestable() => false;

    public FloatingObject(iTextImage image)
    {
        Image = image;
    }
}

/// <summary>
/// DOCX ??? PDF ??
/// </summary>
public class ImageConverter
{
    private readonly WordprocessingDocument _document;
    private readonly ConvertOptions _options;
    private readonly OpenXmlElement? _colorScheme;
    private readonly DrawingRasterizer _drawingRasterizer;

    public ImageConverter(WordprocessingDocument document, ConvertOptions options)
    {
        _document = document;
        _options = options;

        // ???????? (ColorScheme)
        var themePart = _document.MainDocumentPart?.ThemePart;
        _colorScheme = themePart?.Theme?.ThemeElements?.ColorScheme;
        _drawingRasterizer = new DrawingRasterizer(_document, _options);
    }

    /// <summary>
    /// ?????????????????? PDF ??
    /// </summary>
    public List<IElement> ConvertImagesInParagraph(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph, float pageWidth, PdfWriter? writer)
    {
        var elements = new List<IElement>();
        var processedElements = new HashSet<OpenXmlElement>();

        // 0. ?????(TextBox)
        var textBoxes = paragraph.Descendants().Where(e => e.LocalName == "txbxContent" || e.LocalName == "textbox");
        foreach (var txbx in textBoxes)
        {
            var textBoxElement = ConvertTextBox(txbx, pageWidth);
            if (textBoxElement != null)
                elements.Add(textBoxElement);
        }

        // 1. ?? AlternateContent (???????? Choice)
        var alternateContents = paragraph.Descendants<DocumentFormat.OpenXml.AlternateContent>();
        foreach (var ac in alternateContents)
        {
            var choice = ac.Elements<DocumentFormat.OpenXml.AlternateContentChoice>().FirstOrDefault();
            var fallback = ac.Elements<DocumentFormat.OpenXml.AlternateContentFallback>().FirstOrDefault();
            
            // ??? Choice
            var target = choice ?? (OpenXmlElement?)fallback;
            if (target != null)
            {
                var containerImages = ConvertImagesInContainer(target, pageWidth, writer);
                elements.AddRange(containerImages);
                
                // ?????? Drawing ? Picture ????,??????
                foreach (var d in target.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>()) processedElements.Add(d);
                foreach (var p in target.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>()) processedElements.Add(p);
                
                // ??? AlternateContent,???????????????
                // ????????? Descendants ???? Paragraph
            }
            
            // ????,?? AlternateContent ????????????,??????????
            foreach (var d in ac.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>()) processedElements.Add(d);
            foreach (var p in ac.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>()) processedElements.Add(p);
        }

        // 2. ?????? AlternateContent ????? Drawing ? Picture
        var allDrawings = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        foreach (var drawing in allDrawings)
        {
            if (!processedElements.Contains(drawing))
            {
                var image = ExtractImageFromDrawing(drawing, pageWidth);
                ProcessExtractedImage(image, elements, writer);
                processedElements.Add(drawing);
            }
        }

        var allPictures = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>();
        foreach (var picture in allPictures)
        {
            if (!processedElements.Contains(picture))
            {
                var image = ExtractImageFromPicture(picture, pageWidth);
                ProcessExtractedImage(image, elements, writer);
                processedElements.Add(picture);
            }
        }

        return elements;
    }

    /// <summary>
    /// ?????? PDF ??
    /// </summary>
    private IElement? ConvertTextBox(OpenXmlElement textBoxElement, float pageWidth)
    {
        // ?? txbxContent ??
        var txbxContent = textBoxElement.LocalName == "txbxContent" 
            ? textBoxElement 
            : textBoxElement.Descendants().FirstOrDefault(e => e.LocalName == "txbxContent");

        if (txbxContent == null) return null;

        // txbxContent ?? WordprocessingML ?????
        var table = new PdfPTable(1) { WidthPercentage = 100 };
        var cell = new PdfPCell
        {
            Padding = 4f,
            BorderWidth = 0.5f,
            BorderColor = new BaseColor(192, 192, 192)
        };

        // ????:??????
        var textContent = string.Join("\n", txbxContent.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
        if (!string.IsNullOrWhiteSpace(textContent))
        {
            var font = FontFactory.GetFont("Helvetica", 10);
            cell.Phrase = new Phrase(textContent, font);
            table.AddCell(cell);
            return table;
        }

        return null;
    }

    /// <summary>
    /// ???????????
    /// </summary>
    private List<IElement> ConvertImagesInContainer(OpenXmlElement container, float pageWidth, PdfWriter? writer)
    {
        var elements = new List<IElement>();

        var drawings = container.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        foreach (var drawing in drawings)
        {
            var image = ExtractImageFromDrawing(drawing, pageWidth);
            ProcessExtractedImage(image, elements, writer);
        }

        var pictures = container.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>();
        foreach (var picture in pictures)
        {
            var image = ExtractImageFromPicture(picture, pageWidth);
            ProcessExtractedImage(image, elements, writer);
        }

        return elements;
    }

    private static void ProcessExtractedImage(IElement? element, List<IElement> elements, PdfWriter? writer)
    {
        if (element == null) return;
        
        if (element is iTextImage img && img.Alignment == iTextImage.UNDERLYING && writer != null)
        {
            writer.DirectContentUnder.AddImage(img);
        }
        else
        {
            elements.Add(element);
        }
    }

    /// <summary>
    /// ?? Drawing ??(?? IElement,??? Image ? FloatingObject)
    /// </summary>
    private IElement? ExtractImageFromDrawing(DocumentFormat.OpenXml.Wordprocessing.Drawing drawing, float pageWidth)
    {
        // ????
        var inline = drawing.Inline;
        if (inline != null)
        {
            return ExtractImageFromInline(inline, pageWidth);
        }

        // ????
        var anchor = drawing.Anchor;
        if (anchor != null)
        {
            var floatObj = ExtractFloatingObjectFromAnchor(anchor, pageWidth);
            if (floatObj != null) return floatObj;
        }

        return null;
    }

    /// <summary>
    /// ? Inline ????(?????)
    /// </summary>
    private iTextImage? ExtractImageFromInline(DW.Inline inline, float pageWidth)
    {
        var extent = inline.Extent;
        var blipFill = inline.Descendants().FirstOrDefault(e => e.LocalName == "blip");
        var embedId = GetEmbedId(blipFill);
        
        if (embedId != null)
        {
            return CreateImage(embedId, extent, pageWidth, blipFill, isAnchor: false);
        }
        
        return null;
    }

    private static void ProcessExtractedElement(IElement? element, List<IElement> elements, PdfWriter? writer)
    {
        if (element == null) return;
        
        if (element is iTextImage img && img.Alignment == iTextImage.UNDERLYING && writer != null)
        {
            writer.DirectContentUnder.AddImage(img);
        }
        else
        {
            elements.Add(element);
        }
    }

    /// <summary>
    /// ? Anchor ???? (?? FloatingObject)
    /// </summary>
    private FloatingObject? ExtractFloatingObjectFromAnchor(DW.Anchor anchor, float pageWidth)
    {
        var extent = anchor.Extent;
        var blipFill = anchor.Descendants().FirstOrDefault(e => e.LocalName == "blip");
        var embedId = GetEmbedId(blipFill);
        
        iTextImage? image = null;

        if (embedId != null)
        {
            image = CreateImage(embedId, extent, pageWidth, blipFill, isAnchor: true);
        }
        else if (_drawingRasterizer.CanRasterize(anchor))
        {
            var (pxW, pxH) = EstimatePixelSize(extent, pageWidth);
            var png = _drawingRasterizer.RasterizeToPng(anchor, pxW, pxH);
            if (png != null)
            {
                try
                {
                    image = iTextImage.GetInstance(png);
                    if (extent != null)
                    {
                        var widthPt = StyleHelper.EmuToPoints(extent.Cx?.Value ?? 0);
                        var heightPt = StyleHelper.EmuToPoints(extent.Cy?.Value ?? 0);
                        if (widthPt > 0 && heightPt > 0) image.ScaleAbsolute(widthPt, heightPt);
                    }
                }
                catch { }
            }
        }

        if (image == null) return null;

        var floatObj = new FloatingObject(image);
        
        // ??????
        var wrapNone = anchor.GetFirstChild<DW.WrapNone>();
        var wrapSquare = anchor.GetFirstChild<DW.WrapSquare>();
        var wrapTight = anchor.GetFirstChild<DW.WrapTight>();
        var wrapThrough = anchor.GetFirstChild<DW.WrapThrough>();
        var wrapTopBottom = anchor.GetFirstChild<DW.WrapTopBottom>();

        if (wrapSquare != null) floatObj.Wrapping = WrappingStyle.Square;
        else if (wrapTight != null) floatObj.Wrapping = WrappingStyle.Tight;
        else if (wrapThrough != null) floatObj.Wrapping = WrappingStyle.Through;
        else if (wrapTopBottom != null) floatObj.Wrapping = WrappingStyle.TopAndBottom;
        else if (wrapNone != null)
        {
            // WrapNone ??? InFront ? Behind
            // ??? Z-Index ? BehindDoc ??
            // ????:?? BehindDoc=true -> BehindText, ?? InFrontOfText
            floatObj.Wrapping = (anchor.BehindDoc?.Value ?? false) ? WrappingStyle.BehindText : WrappingStyle.InFrontOfText;
        }
        else
        {
            // ?? InFront
            floatObj.Wrapping = WrappingStyle.InFrontOfText;
        }

        // ????
        var posH = anchor.Elements<DW.HorizontalPosition>().FirstOrDefault();
        var posV = anchor.Elements<DW.VerticalPosition>().FirstOrDefault();
        
        if (posH != null && posV != null)
        {
            var offsetXStr = posH.GetFirstChild<DW.PositionOffset>()?.Text;
            var offsetYStr = posV.GetFirstChild<DW.PositionOffset>()?.Text;
            
            float ptX = 0, ptY = 0;
            bool xCalculated = false;
            bool yCalculated = false;

            // ??????
            if (long.TryParse(offsetXStr, out long emuX))
            {
                ptX = StyleHelper.EmuToPoints(emuX);
                var relH = posH.RelativeFrom?.Value;
                if (relH.HasValue)
                {
                    if (relH.Value == DW.HorizontalRelativePositionValues.Margin || relH.Value == DW.HorizontalRelativePositionValues.Page)
                    {
                        if (relH.Value == DW.HorizontalRelativePositionValues.Margin) ptX += _options.MarginLeft;
                        floatObj.PositionIsAbsolute = true;
                    }
                    else if (relH.Value == DW.HorizontalRelativePositionValues.Column)
                    {
                        ptX += _options.MarginLeft; // ??:?????
                        floatObj.PositionIsAbsolute = true;
                    }
                }
                xCalculated = true;
            }
            
            if (long.TryParse(offsetYStr, out long emuY))
            {
                ptY = StyleHelper.EmuToPoints(emuY);
                var relV = posV.RelativeFrom?.Value;
                if (relV.HasValue)
                {
                    if (relV.Value == DW.VerticalRelativePositionValues.Page)
                    {
                        // Page relative Y is from top of page
                        // Convert to iText Y (from bottom) later
                        floatObj.Top = ptY; 
                        floatObj.PositionIsAbsolute = true;
                    }
                    else if (relV.Value == DW.VerticalRelativePositionValues.Margin)
                    {
                        floatObj.Top = ptY + _options.MarginTop;
                        floatObj.PositionIsAbsolute = true;
                    }
                    else if (relV.Value == DW.VerticalRelativePositionValues.Paragraph || relV.Value == DW.VerticalRelativePositionValues.Line)
                    {
                        // ????/?:Top ?????,PositionIsAbsolute = false
                        floatObj.Top = ptY;
                        floatObj.PositionIsAbsolute = false;
                    }
                }
                yCalculated = true;
            }

            // ???????? (Fallback)
            if (!xCalculated)
            {
                var alignH = posH.GetFirstChild<DW.HorizontalAlignment>()?.Text;
                var relH = posH.RelativeFrom?.Value;
                bool isPageRelative = relH.HasValue && relH.Value == DW.HorizontalRelativePositionValues.Page;

                if (alignH == "center") 
                {
                    ptX = (_options.PageSize.Width - image.ScaledWidth) / 2f;
                }
                else if (alignH == "right") 
                {
                    ptX = isPageRelative 
                        ? _options.PageSize.Width - image.ScaledWidth 
                        : _options.PageSize.Width - _options.MarginRight - image.ScaledWidth;
                }
                else 
                {
                    ptX = isPageRelative ? 0 : _options.MarginLeft;
                }
                floatObj.PositionIsAbsolute = true;
            }

            // ???????? (Fallback)
            if (!yCalculated)
            {
                var alignV = posV.GetFirstChild<DW.VerticalAlignment>()?.Text;
                var relV = posV.RelativeFrom?.Value;
                bool isPageRelativeV = relV.HasValue && relV.Value == DW.VerticalRelativePositionValues.Page;

                if (alignV == "center")
                {
                    floatObj.Top = (_options.PageSize.Height - image.ScaledHeight) / 2f;
                }
                else if (alignV == "bottom")
                {
                    floatObj.Top = isPageRelativeV 
                        ? _options.PageSize.Height - image.ScaledHeight
                        : _options.PageSize.Height - _options.MarginBottom - image.ScaledHeight;
                }
                else
                {
                    floatObj.Top = isPageRelativeV ? 0 : _options.MarginTop;
                }
                floatObj.PositionIsAbsolute = true;
            }
            
            floatObj.Left = ptX;
            
            // ?? iText Image ???? (?????????)
            if (floatObj.PositionIsAbsolute)
            {
                var absY = _options.PageSize.Height - floatObj.Top - image.ScaledHeight;
                image.SetAbsolutePosition(floatObj.Left, absY);
            }
        }

        return floatObj;
    }

    private static (int PixelWidth, int PixelHeight) EstimatePixelSize(DW.Extent? extent, float pageWidth)
    {
        // 96 DPI ?,1pt ˜ 96/72 ??
        const float dpi = 96f;
        const float ptToPx = dpi / 72f;

        if (extent?.Cx?.Value is long cx && extent.Cy?.Value is long cy && cx > 0 && cy > 0)
        {
            var widthPt = StyleHelper.EmuToPoints(cx);
            var heightPt = StyleHelper.EmuToPoints(cy);

            if (widthPt <= 0 || heightPt <= 0)
                return ((int)(pageWidth * ptToPx), (int)(pageWidth * 0.6f * ptToPx));

            // ???????
            if (widthPt > pageWidth)
            {
                var ratio = pageWidth / widthPt;
                widthPt = pageWidth;
                heightPt *= ratio;
            }

            var pxW = Math.Max(64, (int)(widthPt * ptToPx));
            var pxH = Math.Max(48, (int)(heightPt * ptToPx));
            return (pxW, pxH);
        }

        // ???????????
        var defaultWidthPt = pageWidth * 0.8f;
        var defaultHeightPt = defaultWidthPt * 0.6f;
        return ((int)(defaultWidthPt * ptToPx), (int)(defaultHeightPt * ptToPx));
    }

    /// <summary>
    /// ??? Picture ????
    /// </summary>
    private iTextImage? ExtractImageFromPicture(DocumentFormat.OpenXml.Wordprocessing.Picture picture, float pageWidth)
    {
        // VML ????
        var imageData = picture.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().FirstOrDefault();
        if (imageData?.RelationshipId?.Value is string relId)
        {
            try
            {
                var mainPart = _document.MainDocumentPart;
                if (mainPart == null) return null;

                var imagePart = mainPart.GetPartById(relId) as ImagePart;
                if (imagePart == null) return null;

                using var stream = imagePart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                var imageBytes = SanitizeImageBytes(ms.ToArray());

                var image = iTextImage.GetInstance(imageBytes);
                image.ScaleToFit(pageWidth, pageWidth); // ????
                image.Alignment = Element.ALIGN_CENTER;
                return image;
            }
            catch
            {
                return null;
            }
        }

        return null;
    }

    /// <summary>
    /// ?? iTextSharp Image
    /// </summary>
    private iTextImage? CreateImage(string embedId, DW.Extent? extent, float pageWidth, OpenXmlElement? blip, bool isAnchor = false)
    {
        try
        {
            var mainPart = _document.MainDocumentPart;
            if (mainPart == null) return null;

            var imagePart = mainPart.GetPartById(embedId) as ImagePart;
            if (imagePart == null) return null;

            using var stream = imagePart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var imageBytes = ApplyBlipEffects(SanitizeImageBytes(ms.ToArray()), blip);
            
            // ????:????
            // ?????????????? 100 ?????
            // ??=100 ????,???????
            if (_options.EnablePerformanceMode && _options.ImageCompressionQuality > 0 && _options.ImageCompressionQuality < 100)
            {
                imageBytes = CompressImage(imageBytes, _options.ImageCompressionQuality);
            }

            var image = iTextImage.GetInstance(imageBytes);

            // ??????
            if (extent != null)
            {
                var widthPt = StyleHelper.EmuToPoints(extent.Cx?.Value ?? 0);
                var heightPt = StyleHelper.EmuToPoints(extent.Cy?.Value ?? 0);

                if (widthPt > 0 && heightPt > 0)
                {
                    if (!isAnchor)
                    {
                        // ???????
                        if (widthPt > pageWidth)
                        {
                            var ratio = pageWidth / widthPt;
                            widthPt = pageWidth;
                            heightPt *= ratio;
                        }
                    }

                    image.ScaleAbsolute(widthPt, heightPt);
                }
                else
                {
                    if (!isAnchor) image.ScaleToFit(pageWidth, pageWidth);
                }
            }
            else
            {
                if (!isAnchor) image.ScaleToFit(pageWidth, pageWidth);
            }

            if (!isAnchor) image.Alignment = Element.ALIGN_CENTER;
            return image;
        }
        catch
        {
            return null;
        }
    }

    private byte[] ApplyBlipEffects(byte[] imageBytes, OpenXmlElement? blip)
    {
        if (blip == null) return imageBytes;

        var duotone = blip.ChildElements.FirstOrDefault(e => e.LocalName == "duotone");
        if (duotone != null)
        {
            var colorNodes = duotone.Descendants()
                .Where(e => e.LocalName == "schemeClr" || e.LocalName == "srgbClr" || e.LocalName == "prstClr")
                .Take(2)
                .ToList();

            if (colorNodes.Count == 2)
            {
                var c1 = ResolveColor(colorNodes[0]);
                var c2 = ResolveColor(colorNodes[1]);
                if (c1 != null && c2 != null)
                {
                    // BakeDuotone has been disabled: it mathematically destroyed the transparent gradients.
                    // Instead, we use the original image which is already correctly tinted but transparent,
                    // and let PdfEngine composite it onto a white background to achieve the correct light blue visual.
                    return imageBytes;
                }
            }
        }

        return imageBytes;
    }

    private byte[]? BakeDuotone(byte[] imageBytes, BaseColor color1, BaseColor color2)
    {
        try
        {
            using var srcStream = new SKMemoryStream(imageBytes);
            using var codec = SKCodec.Create(srcStream);
            if (codec == null) return null;

            var srcInfo = codec.Info;
            var dstInfo = new SKImageInfo(srcInfo.Width, srcInfo.Height, SKColorType.Bgra8888, SKAlphaType.Unpremul, SKColorSpace.CreateSrgb());
            using var bitmap = new SKBitmap(dstInfo);
            var result = codec.GetPixels(dstInfo, bitmap.GetPixels());
            if (result != SKCodecResult.Success && result != SKCodecResult.IncompleteInput) return null;

            var cA = new SKColor((byte)color1.R, (byte)color1.G, (byte)color1.B, 255);
            var cB = new SKColor((byte)color2.R, (byte)color2.G, (byte)color2.B, 255);

            var lumA = (0.2126f * cA.Red + 0.7152f * cA.Green + 0.0722f * cA.Blue) / 255f;
            var lumB = (0.2126f * cB.Red + 0.7152f * cB.Green + 0.0722f * cB.Blue) / 255f;

            var dark = lumA <= lumB ? cA : cB;
            var light = lumA <= lumB ? cB : cA;

            var pixels = bitmap.GetPixelSpan();
            for (var i = 0; i + 3 < pixels.Length; i += 4)
            {
                var b0 = pixels[i];
                var g0 = pixels[i + 1];
                var r0 = pixels[i + 2];
                var a0 = pixels[i + 3];

                if (a0 == 0) continue;

                var l = (0.2126f * r0 + 0.7152f * g0 + 0.0722f * b0) / 255f;

                var r = (byte)Math.Clamp(dark.Red + (light.Red - dark.Red) * l, 0, 255);
                var g = (byte)Math.Clamp(dark.Green + (light.Green - dark.Green) * l, 0, 255);
                var b = (byte)Math.Clamp(dark.Blue + (light.Blue - dark.Blue) * l, 0, 255);

                pixels[i] = b;
                pixels[i + 1] = g;
                pixels[i + 2] = r;
                pixels[i + 3] = a0;
            }

            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            return data.ToArray();
        }
        catch
        {
            return null;
        }
    }

    private BaseColor? ResolveColor(OpenXmlElement node)
    {
        var val = node.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (string.IsNullOrWhiteSpace(val)) return null;

        BaseColor? baseColor = null;

        if (node.LocalName == "schemeClr")
        {
            baseColor = StyleHelper.ResolveSchemeColor(_colorScheme, val);
        }
        else if (node.LocalName == "srgbClr")
        {
            baseColor = StyleHelper.HexToBaseColor(val);
        }
        else if (node.LocalName == "prstClr")
        {
            baseColor = val.ToLowerInvariant() switch
            {
                "black" => new BaseColor(0, 0, 0),
                "white" => new BaseColor(255, 255, 255),
                "red" => new BaseColor(255, 0, 0),
                "green" => new BaseColor(0, 128, 0),
                "blue" => new BaseColor(0, 0, 255),
                "yellow" => new BaseColor(255, 255, 0),
                "cyan" => new BaseColor(0, 255, 255),
                "magenta" => new BaseColor(255, 0, 255),
                "gray" or "grey" => new BaseColor(128, 128, 128),
                _ => null
            };
        }

        if (baseColor != null)
        {
            return ApplyDrawingColorModifiers(baseColor, node);
        }

        return null;
    }

    private static BaseColor ApplyDrawingColorModifiers(BaseColor color, OpenXmlElement colorNode)
    {
        // ??????? (? DrawingML ?,? a:tint, a:shade, a:lumMod, a:lumOff)
        // ????????????? (1/1000 of a percent) ???,? 100000 = 100%

        int r = color.R;
        int g = color.G;
        int b = color.B;

        var children = colorNode.ChildElements;
        
        foreach (var modifier in children)
        {
            var valStr = modifier.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
            if (string.IsNullOrEmpty(valStr) || !int.TryParse(valStr, out int valEmu)) continue;

            double ratio = valEmu / 100000.0;

            switch (modifier.LocalName)
            {
                case "tint": // a:tint (????)
                    r = (int)Math.Round(r + (255 - r) * (1 - ratio));
                    g = (int)Math.Round(g + (255 - g) * (1 - ratio));
                    b = (int)Math.Round(b + (255 - b) * (1 - ratio));
                    break;
                case "shade": // a:shade (????,???)
                    r = (int)Math.Round(r * ratio);
                    g = (int)Math.Round(g * ratio);
                    b = (int)Math.Round(b * ratio);
                    break;
                case "lumMod": // a:lumMod (????)
                    // HSL ???????,?????????RGB
                    r = (int)Math.Round(r * ratio);
                    g = (int)Math.Round(g * ratio);
                    b = (int)Math.Round(b * ratio);
                    break;
                case "lumOff": // a:lumOff (????)
                    int offset = (int)Math.Round(255 * ratio);
                    r += offset;
                    g += offset;
                    b += offset;
                    break;
                case "alpha":
                    // ???????
                    break;
            }
        }

        return new BaseColor(
            Math.Clamp(r, 0, 255),
            Math.Clamp(g, 0, 255),
            Math.Clamp(b, 0, 255)
        );
    }

    private static byte[] SanitizeImageBytes(byte[] imageBytes)
    {
        if (IsPng(imageBytes))
        {
            return StripPngChunk(imageBytes, "iCCP");
        }

        return imageBytes;
    }

    private static bool IsPng(byte[] bytes)
    {
        return bytes.Length >= 8
               && bytes[0] == 0x89
               && bytes[1] == 0x50
               && bytes[2] == 0x4E
               && bytes[3] == 0x47
               && bytes[4] == 0x0D
               && bytes[5] == 0x0A
               && bytes[6] == 0x1A
               && bytes[7] == 0x0A;
    }

    private static byte[] StripPngChunk(byte[] pngBytes, string chunkType)
    {
        if (!IsPng(pngBytes) || chunkType.Length != 4) return pngBytes;

        try
        {
            var changed = false;
            using var output = new MemoryStream(pngBytes.Length);

            output.Write(pngBytes, 0, 8);

            var offset = 8;
            while (offset + 8 <= pngBytes.Length)
            {
                var length = ReadInt32BigEndian(pngBytes, offset);
                if (length < 0) return pngBytes;

                if (offset + 12L + length > pngBytes.Length) return pngBytes;

                var type = Encoding.ASCII.GetString(pngBytes, offset + 4, 4);
                var total = 12 + length;

                if (string.Equals(type, chunkType, StringComparison.Ordinal))
                {
                    changed = true;
                }
                else
                {
                    output.Write(pngBytes, offset, total);
                }

                offset += total;

                if (string.Equals(type, "IEND", StringComparison.Ordinal)) break;
            }

            return changed ? output.ToArray() : pngBytes;
        }
        catch
        {
            return pngBytes;
        }
    }

    private static int ReadInt32BigEndian(byte[] bytes, int offset)
    {
        return (bytes[offset] << 24)
               | (bytes[offset + 1] << 16)
               | (bytes[offset + 2] << 8)
               | bytes[offset + 3];
    }
    
    /// <summary>
    /// ???????????
    /// </summary>
    private static byte[] CompressImage(byte[] imageBytes, int quality)
    {
        try
        {
            // ???? SkiaSharp ????
            using var srcStream = new SKMemoryStream(imageBytes);
            using var codec = SKCodec.Create(srcStream);
            if (codec == null) return imageBytes;
            
            var srcInfo = codec.Info;
            var dstInfo = new SKImageInfo(srcInfo.Width, srcInfo.Height, SKColorType.Bgra8888, SKAlphaType.Premul);
            
            using (var bitmap = new SKBitmap(dstInfo))
            {
                var result = codec.GetPixels(dstInfo, bitmap.GetPixels());
                if (result != SKCodecResult.Success && result != SKCodecResult.IncompleteInput) return imageBytes;
                
                using var image = SKImage.FromBitmap(bitmap);
                // ??????????
                var format = quality >= 90 ? SKEncodedImageFormat.Png : SKEncodedImageFormat.Jpeg;
                using var data = image.Encode(format, Math.Clamp(quality, 1, 100));
                return data.ToArray();
            }
        }
        catch
        {
            // ????,????
            return imageBytes;
        }
    }
    /// <summary>
    /// ?? Blip ??????(? Duotone / ColorChange)
    /// </summary>
    private BaseColor? GetBlipEffectColor(OpenXmlElement? blip)
    {
        if (blip == null) return null;

        // ?? duotone
        var duotone = blip.ChildElements.FirstOrDefault(e => e.LocalName == "duotone");
        if (duotone != null)
        {
            // ?? duotone ????? (schemeClr, srgbClr, prstClr)
            // ?? Word ????? duotone ????????
            var clrNode = duotone.Descendants().LastOrDefault(e => e.LocalName == "schemeClr" || e.LocalName == "srgbClr");
            if (clrNode != null)
            {
                if (clrNode.LocalName == "schemeClr")
                {
                    string? val = null;
                    foreach (var attr in clrNode.GetAttributes())
                    {
                        if (attr.LocalName == "val") { val = attr.Value; break; }
                    }
                    return StyleHelper.ResolveSchemeColor(_colorScheme, val);
                }
                else if (clrNode.LocalName == "srgbClr")
                {
                    string? val = null;
                    foreach (var attr in clrNode.GetAttributes())
                    {
                        if (attr.LocalName == "val") { val = attr.Value; break; }
                    }
                    return StyleHelper.HexToBaseColor(val);
                }
            }
        }

        // ?? clrChange
        var clrChange = blip.ChildElements.FirstOrDefault(e => e.LocalName == "clrChange");
        if (clrChange != null)
        {
            var toClr = clrChange.Descendants().FirstOrDefault(e => e.LocalName == "toClr");
            var clrNode = toClr?.Descendants().FirstOrDefault(e => e.LocalName == "schemeClr" || e.LocalName == "srgbClr");
            if (clrNode != null)
            {
                string? val = null;
                foreach (var attr in clrNode.GetAttributes())
                {
                    if (attr.LocalName == "val") { val = attr.Value; break; }
                }

                if (clrNode.LocalName == "schemeClr")
                    return StyleHelper.ResolveSchemeColor(_colorScheme, val);
                else
                    return StyleHelper.HexToBaseColor(val);
            }
        }

        return null;
    }

    /// <summary>
    /// ? blip ????? r:embed ???
    /// </summary>
    private static string? GetEmbedId(OpenXmlElement? blip)
    {
        if (blip == null) return null;
        foreach (var attr in blip.GetAttributes())
        {
            if (attr.LocalName == "embed") return attr.Value;
        }
        return null;
    }
}
