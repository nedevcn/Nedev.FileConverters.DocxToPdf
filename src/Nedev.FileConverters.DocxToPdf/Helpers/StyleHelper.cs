using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.PdfEngine;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

/// <summary>
/// DOCX ??? PDF ?????????
/// </summary>
public static class StyleHelper
{
    /// <summary>
    /// ? DOCX ??(hex ???)?? iTextSharp BaseColor
    /// </summary>
    public static BaseColor? HexToBaseColor(string? hexColor)
    {
        if (string.IsNullOrWhiteSpace(hexColor) || hexColor == "auto")
            return null;

        hexColor = hexColor.TrimStart('#');

        if (hexColor.Length == 6 &&
            int.TryParse(hexColor[..2], NumberStyles.HexNumber, null, out var r) &&
            int.TryParse(hexColor[2..4], NumberStyles.HexNumber, null, out var g) &&
            int.TryParse(hexColor[4..6], NumberStyles.HexNumber, null, out var b))
        {
            return new BaseColor(r, g, b);
        }

        return null;
    }

    /// <summary>
    /// ? DOCX ?????? iTextSharp ????
    /// </summary>
    public static int ToiTextAlignment(JustificationValues? justification)
    {
        if (justification == null) return Element.ALIGN_LEFT;
        
        if (justification.Equals(JustificationValues.Center))
            return Element.ALIGN_CENTER;
        if (justification.Equals(JustificationValues.Right))
            return Element.ALIGN_RIGHT;
        if (justification.Equals(JustificationValues.Both) || justification.Equals(JustificationValues.Distribute))
            return Element.ALIGN_JUSTIFIED;
            
        return Element.ALIGN_LEFT;
    }

    /// <summary>
    /// ? DOCX Twips(1/20 pt)?? PDF points
    /// </summary>
    public static float TwipsToPoints(string? twipsStr)
    {
        if (string.IsNullOrWhiteSpace(twipsStr)) return 0;
        if (float.TryParse(twipsStr, out var twips))
            return twips / 20f;
        return 0;
    }

    /// <summary>
    /// ? DOCX Twips(1/20 pt)?? PDF points
    /// </summary>
    public static float TwipsToPoints(int twips) => twips / 20f;

    /// <summary>
    /// ? EMU(English Metric Units)?? PDF points
    /// </summary>
    public static float EmuToPoints(long emu) => emu / 914400f * 72f;

    /// <summary>
    /// ??????????
    /// </summary>
    public static float GetHeadingFontSize(string? styleId)
    {
        return styleId?.ToLowerInvariant() switch
        {
            "heading1" or "1" => 24f,
            "heading2" or "2" => 20f,
            "heading3" or "3" => 16f,
            "heading4" or "4" => 14f,
            "heading5" or "5" => 13f,
            "heading6" or "6" => 12f,
            _ => 0f
        };
    }

    /// <summary>
    /// ???????????
    /// </summary>
    public static bool IsHeadingStyle(string? styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId)) return false;
        return styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
               || styleId.StartsWith("heading", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// ? DXA(1/20 pt)?? points,???????
    /// </summary>
    public static float DxaToPoints(string? dxaStr)
    {
        if (string.IsNullOrWhiteSpace(dxaStr)) return 0;
        if (float.TryParse(dxaStr, out var dxa))
            return dxa / 20f;
        return 0;
    }

    /// <summary>
    /// ? DOCX half-point ?? point(???)
    /// </summary>
    public static float HalfPointToPoint(string? halfPtStr)
    {
        if (string.IsNullOrWhiteSpace(halfPtStr)) return 0;
        if (float.TryParse(halfPtStr, out var halfPt))
            return halfPt / 2f;
        return 0;
    }

    /// <summary>
    /// ? DOCX ?????? iTextSharp ????(???)
    /// </summary>
    public static float GetBorderWidth(BorderValues? borderType)
    {
        if (borderType == null || borderType.Equals(BorderValues.None) || borderType.Equals(BorderValues.Nil))
            return 0f;
            
        if (borderType.Equals(BorderValues.Single)) return 0.5f;
        if (borderType.Equals(BorderValues.Thick)) return 1.5f;
        if (borderType.Equals(BorderValues.Double)) return 1f;
        
        return 0.5f;
    }

    public static float GetBorderWidth(BorderType? border)
    {
        if (border == null) return 0f;
        var val = border.Val?.Value;
        if (val != null && (val.Equals(BorderValues.None) || val.Equals(BorderValues.Nil))) return 0f;

        if (border.Size?.Value is uint sz && sz > 0)
            return sz / 8f;

        return GetBorderWidth(val);
    }

    public static BaseColor? ResolveBorderColor(OpenXmlElement? colorScheme, BorderType? border)
    {
        if (border == null) return null;

        var directHex = border.Color?.Value;
        var direct = HexToBaseColor(directHex);
        if (direct != null)
            return ApplyTintShade(direct, border.ThemeTint?.Value, border.ThemeShade?.Value);

        var themeKey = MapThemeColorName(border.ThemeColor?.Value.ToString());
        var theme = ResolveSchemeColor(colorScheme, themeKey);
        if (theme != null)
            return ApplyTintShade(theme, border.ThemeTint?.Value, border.ThemeShade?.Value);

        return null;
    }

    public static BaseColor? ResolveWordColor(OpenXmlElement? colorScheme, DocumentFormat.OpenXml.Wordprocessing.Color? colorNode)
    {
        if (colorNode == null) return null;

        var directHex = colorNode.Val?.Value;
        var direct = HexToBaseColor(directHex);
        if (direct != null)
            return ApplyTintShade(direct, colorNode.ThemeTint?.Value, colorNode.ThemeShade?.Value);

        var themeKey = MapThemeColorName(colorNode.ThemeColor?.Value.ToString());
        var theme = ResolveSchemeColor(colorScheme, themeKey);
        if (theme != null)
            return ApplyTintShade(theme, colorNode.ThemeTint?.Value, colorNode.ThemeShade?.Value);

        return null;
    }

    public static BaseColor? ResolveShadingFill(OpenXmlElement? colorScheme, Shading? shading)
    {
        if (shading == null) return null;

        var directHex = shading.Fill?.Value;
        var direct = HexToBaseColor(directHex);
        if (direct != null)
            return ApplyTintShade(direct, shading.ThemeFillTint?.Value, shading.ThemeFillShade?.Value);

        var themeKey = MapThemeColorName(shading.ThemeFill?.Value.ToString());
        var theme = ResolveSchemeColor(colorScheme, themeKey);
        if (theme != null)
            return ApplyTintShade(theme, shading.ThemeFillTint?.Value, shading.ThemeFillShade?.Value);

        return null;
    }

    private static string? MapThemeColorName(string? themeColor)
    {
        if (string.IsNullOrWhiteSpace(themeColor)) return null;
        var k = themeColor.Trim().ToLowerInvariant();

        return k switch
        {
            "text1" => "dk1",
            "background1" => "lt1",
            "text2" => "dk2",
            "background2" => "lt2",
            "dark1" => "dk1",
            "light1" => "lt1",
            "dark2" => "dk2",
            "light2" => "lt2",
            "accent1" => "accent1",
            "accent2" => "accent2",
            "accent3" => "accent3",
            "accent4" => "accent4",
            "accent5" => "accent5",
            "accent6" => "accent6",
            "hyperlink" => "hlink",
            "followedhyperlink" => "folHlink",
            _ => k
        };
    }

    private static BaseColor ApplyTintShade(BaseColor color, string? tintHex, string? shadeHex)
    {
        var r = color.R;
        var g = color.G;
        var b = color.B;

        if (!string.IsNullOrWhiteSpace(tintHex) && int.TryParse(tintHex, NumberStyles.HexNumber, null, out var tint))
        {
            r = (byte)ApplyTintComponent(r, tint);
            g = (byte)ApplyTintComponent(g, tint);
            b = (byte)ApplyTintComponent(b, tint);
        }

        if (!string.IsNullOrWhiteSpace(shadeHex) && int.TryParse(shadeHex, NumberStyles.HexNumber, null, out var shade))
        {
            r = (byte)ApplyShadeComponent(r, shade);
            g = (byte)ApplyShadeComponent(g, shade);
            b = (byte)ApplyShadeComponent(b, shade);
        }

        return new BaseColor((int)r, (int)g, (int)b);
    }

    private static int ApplyTintComponent(int c, int tint)
    {
        tint = Math.Clamp(tint, 0, 255);
        var v = c + (255 - c) * (tint / 255f);
        return (int)Math.Round(Math.Clamp(v, 0, 255));
    }

    private static int ApplyShadeComponent(int c, int shade)
    {
        shade = Math.Clamp(shade, 0, 255);
        var v = c * (shade / 255f);
        return (int)Math.Round(Math.Clamp(v, 0, 255));
    }
    /// <summary>
    /// ??????
    /// </summary>
    public static BaseColor? ResolveSchemeColor(OpenXmlElement? colorScheme, string? schemeColorName)
    {
        if (colorScheme == null || string.IsNullOrEmpty(schemeColorName)) return null;

        // ?????????(? <a:accent1>)
        var colorElement = colorScheme.Elements().FirstOrDefault(e => e.LocalName.Equals(schemeColorName, StringComparison.OrdinalIgnoreCase));
        if (colorElement == null) return null;

        // ?? srgbClr (RgbColorModelHex)
        var srgb = colorElement.Descendants().FirstOrDefault(e => e.LocalName == "srgbClr" || e.LocalName == "rgbColorModelHex");
        if (srgb != null)
        {
            foreach (var attr in srgb.GetAttributes())
            {
                if (attr.LocalName == "val") return HexToBaseColor(attr.Value);
            }
        }

        // ?? sysClr (SystemColor)
        var sys = colorElement.Descendants().FirstOrDefault(e => e.LocalName == "sysClr" || e.LocalName == "systemColor");
        if (sys != null)
        {
            foreach (var attr in sys.GetAttributes())
            {
                if (attr.LocalName == "lastClr") return HexToBaseColor(attr.Value);
            }
        }

        return null;
    }

    /// <summary>
    /// ?????????(?????????????)
    /// </summary>
    public static bool IsDarkColor(BaseColor? color)
    {
        if (color == null) return false;
        // ???????: (0.299*R + 0.587*G + 0.114*B)
        double luminance = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255.0;
        return luminance < 0.5;
    }
}
