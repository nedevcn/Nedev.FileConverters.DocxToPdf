using Nedev.DocxToPdf;
using Nedev.DocxToPdf.Helpers;
using Nedev.DocxToPdf.PdfEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// 如果第一个参数是 --test，运行PDF引擎测试
/* if (args.Length >= 1 && args[0] == "--test")
{
    var test = new PdfEngineTest();
    test.RunAllTests();
    return;
} */

var docxPath = args.Length >= 1 ? args[0] : "test.docx";
var pdfPath = args.Length >= 2 ? args[1] : "test.pdf";

docxPath = Path.GetFullPath(docxPath);
pdfPath = Path.GetFullPath(pdfPath);

DumpImages(docxPath);
new DocxToPdfConverter().Convert(docxPath, pdfPath);
Console.WriteLine(pdfPath);

static void DumpImages(string docxPath)
{
    using var doc = WordprocessingDocument.Open(docxPath, false);
    var mainPart = doc.MainDocumentPart;
    var body = mainPart?.Document?.Body;
    if (mainPart == null || body == null) return;

    var colorScheme = mainPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;

    var drawings = body.Descendants<Drawing>().Take(25).ToList();
    Console.WriteLine($"Drawings={drawings.Count}");

    for (var i = 0; i < drawings.Count; i++)
    {
        var drawing = drawings[i];
        var anchor = drawing.Anchor;
        var behindDoc = anchor?.BehindDoc?.Value ?? false;

        var blip = drawing.Descendants().FirstOrDefault(e => e.LocalName == "blip");
        var embedId = GetAttr(blip, "embed");
        var hasDuotone = blip?.ChildElements.Any(e => e.LocalName == "duotone") ?? false;
        var hasClrChange = blip?.ChildElements.Any(e => e.LocalName == "clrChange") ?? false;
        var effectColor = GetBlipEffectColor(blip, colorScheme);

        var partUri = "";
        if (!string.IsNullOrWhiteSpace(embedId))
        {
            if (mainPart.GetPartById(embedId) is ImagePart imagePart)
            {
                partUri = imagePart.Uri.ToString();
            }
        }

        var effectRgb = effectColor == null ? "" : $"{effectColor.R},{effectColor.G},{effectColor.B}";
        Console.WriteLine($"{i:00} behind={behindDoc} embed={embedId} part={partUri} duotone={hasDuotone} clrChange={hasClrChange} effect={effectRgb}");
    }
}

static string? GetAttr(OpenXmlElement? element, string localName)
{
    if (element == null) return null;
    foreach (var attr in element.GetAttributes())
    {
        if (attr.LocalName == localName) return attr.Value;
    }
    return null;
}

static BaseColor? GetBlipEffectColor(OpenXmlElement? blip, OpenXmlElement? colorScheme)
{
    if (blip == null) return null;

    var duotone = blip.ChildElements.FirstOrDefault(e => e.LocalName == "duotone");
    if (duotone != null)
    {
        var clrNode = duotone.Descendants().LastOrDefault(e => e.LocalName == "schemeClr" || e.LocalName == "srgbClr");
        if (clrNode != null)
        {
            var val = GetAttr(clrNode, "val");
            if (clrNode.LocalName == "schemeClr") return StyleHelper.ResolveSchemeColor(colorScheme, val);
            if (clrNode.LocalName == "srgbClr") return StyleHelper.HexToBaseColor(val);
        }
    }

    var clrChange = blip.ChildElements.FirstOrDefault(e => e.LocalName == "clrChange");
    if (clrChange != null)
    {
        var toClr = clrChange.Descendants().FirstOrDefault(e => e.LocalName == "toClr");
        var clrNode = toClr?.Descendants().FirstOrDefault(e => e.LocalName == "schemeClr" || e.LocalName == "srgbClr");
        if (clrNode != null)
        {
            var val = GetAttr(clrNode, "val");
            if (clrNode.LocalName == "schemeClr") return StyleHelper.ResolveSchemeColor(colorScheme, val);
            if (clrNode.LocalName == "srgbClr") return StyleHelper.HexToBaseColor(val);
        }
    }

    return null;
}
