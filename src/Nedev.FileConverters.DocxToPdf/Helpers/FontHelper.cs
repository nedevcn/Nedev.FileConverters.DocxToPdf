using System.Collections.Concurrent;
using DocumentFormat.OpenXml.Wordprocessing;
using Nedev.FileConverters.DocxToPdf.Models;
using Nedev.FileConverters.DocxToPdf.PdfEngine;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;

namespace Nedev.FileConverters.DocxToPdf.Helpers;

/// <summary>
/// ???????,???????????
/// </summary>
public class FontHelper
{
    private readonly ConvertOptions _options;
    private readonly DocumentFormat.OpenXml.OpenXmlElement? _colorScheme;
    private readonly ConcurrentDictionary<string, iTextFont> _fontCache = new();
    // map common Chinese font names (as they appear in DOCX) to their
    // English equivalents that iText/FontFactory understands.  Keys are encoded
    // with \u escapes so the source file remains plain ASCII and avoids
    // encoding problems on non‑UTF8 systems.
    private static readonly Dictionary<string, string> _fontNameMap = new(StringComparer.OrdinalIgnoreCase)
    {
        { "\u5FAE\u8F6F\u96C5\u9ED1", "Microsoft YaHei" }, // 微软雅黑
        { "\u5B8B\u4F53", "SimSun" },                       // 宋体
        { "\u65B0\u5B8B\u4F53", "NSimSun" },               // 新宋体
        { "\u9ED1\u4F53", "SimHei" },                       // 黑体
        { "\u6977\u4F53", "KaiTi" },                       // 楷体
        { "\u4F3C\u5B8B", "FangSong" },                    // 仿宋
        { "\u7B49\u7EBF", "DengXian" },                    // 等线
        { "Microsoft YaHei UI", "Microsoft YaHei" },
        { "Microsoft Yahei", "Microsoft YaHei" }
    };
    private bool _fontsRegistered;

    public FontHelper(ConvertOptions options, DocumentFormat.OpenXml.OpenXmlElement? colorScheme = null)
    {
        _options = options;
        _colorScheme = colorScheme;
    }

    /// <summary>
    /// ??????????????
    /// </summary>
    public void RegisterFonts()
    {
        if (_fontsRegistered) return;

        // ????:??????(???????????)
        if (_options.SkipUnusedFonts)
        {
            // ?????????
            _fontsRegistered = true;
            return;
        }

        // ????????
        // 初始化系统字体映射
        SystemFontProvider.Initialize();

        if (OperatingSystem.IsWindows())
        {
            var winFontDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
            if (Directory.Exists(winFontDir))
                FontFactory.RegisterDirectory(winFontDir);
        }
        else if (OperatingSystem.IsLinux())
        {
            var linuxFontDirs = new[] { "/usr/share/fonts", "/usr/local/share/fonts" };
            foreach (var dir in linuxFontDirs)
            {
                if (Directory.Exists(dir))
                    FontFactory.RegisterDirectory(dir);
            }
        }
        else if (OperatingSystem.IsMacOS())
        {
            var macFontDirs = new[] { "/System/Library/Fonts", "/Library/Fonts" };
            foreach (var dir in macFontDirs)
            {
                if (Directory.Exists(dir))
                    FontFactory.RegisterDirectory(dir);
            }
        }

        // ????????
        foreach (var dir in _options.ExtraFontDirectories)
        {
            if (Directory.Exists(dir))
                FontFactory.RegisterDirectory(dir);
        }

        // ?????????????(????)
        if (OperatingSystem.IsWindows())
        {
            var winFontDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
            var boldFonts = new[] { "msyhbd.ttc,0", "simhei.ttf", "msyh.ttc,0" };
            foreach (var bf in boldFonts)
            {
                var path = Path.Combine(winFontDir, bf.Split(',')[0]);
                if (File.Exists(path)) FontFactory.Register(Path.Combine(winFontDir, bf));
            }
        }

        _fontsRegistered = true;
    }

    /// <summary>
    /// 获取默认中文字体（带回退机制）
    /// </summary>
    public string GetDefaultFontName()
    {
        // 尝试常见的跨平台中文字体名称
        var chineseFontNames = new[]
        {
            "Microsoft YaHei", "SimSun", "NSimSun", 
            "SimHei", "KaiTi", "FangSong", "STFangsong", "STSong",
            "PingFang SC", "Hiragino Sans GB", "WenQuanYi Micro Hei"
        };

        foreach (var fontName in chineseFontNames)
        {
            // 优先检查 FontFactory 注册的
            try
            {
                if (FontFactory.IsRegistered(fontName))
                    return fontName;
            }
            catch { }
            
            // 其次检查系统映射表
            if (SystemFontProvider.GetFontPath(fontName) != null)
                return fontName;
        }

        // 终极 fallback：默认使用 Helvetica（仅英文）或第一个可用的中文字体
        return "Helvetica";
    }
    /// <summary>
    /// ?? DOCX RunProperties ?? iTextSharp Font
    /// </summary>
    public iTextFont GetFont(RunProperties? runProperties, float? defaultSize = null, bool boldOverride = false)
    {
        return GetFont(runProperties, null, defaultSize, boldOverride);
    }

    /// <summary>
    /// ?? DOCX RunProperties ? ParagraphMarkRunProperties ?? iTextSharp Font
    /// </summary>
    public iTextFont GetFont(RunProperties? runProperties, ParagraphMarkRunProperties? paraRunProps, float? defaultSize = null, bool boldOverride = false, string? preferredFontName = null)
    {
        var fontSize = defaultSize ?? _options.DefaultFontSize;
        var fontStyle = iTextFont.NORMAL;
        BaseColor? color = null;
        string? fontName = null;

        // 优先使用传入的首选字体名称（来自样式继承）
        if (!string.IsNullOrWhiteSpace(preferredFontName))
        {
            fontName = preferredFontName;
            if (_fontNameMap.TryGetValue(fontName, out var mappedName))
                fontName = mappedName;
        }

        // 如果没有首选字体，则从 Run 属性中获取
        if (string.IsNullOrWhiteSpace(fontName))
        {
            var rFonts = runProperties?.GetFirstChild<RunFonts>() ?? paraRunProps?.GetFirstChild<RunFonts>();
            fontName = rFonts?.Ascii?.Value ?? rFonts?.EastAsia?.Value ?? rFonts?.HighAnsi?.Value;

            // ????
            if (fontName != null && _fontNameMap.TryGetValue(fontName, out var englishName))
                fontName = englishName;
        }

        var sz = runProperties?.GetFirstChild<FontSize>() ?? paraRunProps?.GetFirstChild<FontSize>();
        if (sz?.Val?.Value is string sizeStr && float.TryParse(sizeStr, out var halfPt))
            fontSize = halfPt / 2f;

        var bold = runProperties?.GetFirstChild<Bold>() ?? paraRunProps?.GetFirstChild<Bold>();
        if (boldOverride || (bold != null && (bold.Val == null || bold.Val.Value)))
            fontStyle |= iTextFont.BOLD;

        var italic = runProperties?.GetFirstChild<Italic>() ?? paraRunProps?.GetFirstChild<Italic>();
        if (italic != null && (italic.Val == null || italic.Val.Value))
            fontStyle |= iTextFont.ITALIC;

        var underline = runProperties?.GetFirstChild<Underline>() ?? paraRunProps?.GetFirstChild<Underline>();
        if (underline?.Val != null && underline.Val != UnderlineValues.None)
            fontStyle |= iTextFont.UNDERLINE;

        var strike = runProperties?.GetFirstChild<Strike>() ?? paraRunProps?.GetFirstChild<Strike>();
        if (strike != null && (strike.Val == null || strike.Val.Value))
            fontStyle |= iTextFont.STRIKETHRU;

        var colorNode = runProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>() ?? paraRunProps?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>();
        if (colorNode != null)
        {
            if (string.Equals(colorNode.Val?.Value, "auto", StringComparison.OrdinalIgnoreCase))
            {
                color = BaseColor.Black;
            }
            else
            {
                color = StyleHelper.ResolveWordColor(_colorScheme, colorNode);
            }
        }


        // ??? - ???? boldOverride,????????????????????????
        var cacheKey = $"{fontName ?? "default"}_{fontSize}_{fontStyle}_{color?.ToArgb() ?? 0}_{boldOverride}";

        return _fontCache.GetOrAdd(cacheKey, _ =>
        {
            var wantsBold = boldOverride || (fontStyle & iTextFont.BOLD) != 0;

            // ????????
            if (!string.IsNullOrWhiteSpace(fontName))
            {
                try
                {
                    if (wantsBold)
                    {
                        var trueBold = TryGetTrueBoldFont(fontName, fontSize, fontStyle, color);
                        if (trueBold != null) return trueBold;
                    }

                    if (FontFactory.IsRegistered(fontName))
                    {
                        var regFont = FontFactory.GetFont(fontName, fontSize, fontStyle, color ?? BaseColor.Black);
                        if (regFont.BaseFont != null)
                            return NormalizeBoldStyleIfNeeded(regFont, wantsBold, fontSize, fontStyle, color);
                    }
                }
                catch
                {
                    // ??
                }
            }

            // ???????
            var defaultFontName = GetDefaultFontName();
            
            // ?????,?????????,???????(SimHei),????????????
            if (wantsBold)
            {
                if (fontName != "SimHei" && fontName != "??")
                {
                    if (FontFactory.IsRegistered("SimHei"))
                    {
                        var heiFont = FontFactory.GetFont("SimHei", fontSize, fontStyle | iTextFont.BOLD, color ?? BaseColor.Black);
                        if (heiFont.BaseFont != null) return NormalizeBoldStyleIfNeeded(heiFont, wantsBold, fontSize, fontStyle, color);
                    }
                }
            }

            return new iTextFont(defaultFontName, fontSize, fontStyle, color ?? BaseColor.Black);
        });
    }

    /// <summary>
    /// Generate a font from DrawingML <see cref="DocumentFormat.OpenXml.Drawing.RunProperties"/>.
    /// This mirrors the logic used for Wordprocessing run properties but handles the
    /// different XML classes (latinFont, RgbColorModelHex, etc.).
    /// </summary>
    public iTextFont GetFont(DocumentFormat.OpenXml.Drawing.RunProperties? runPr, float? defaultSize = null)
    {
        var fontSize = defaultSize ?? _options.DefaultFontSize;
        int fontStyle = iTextFont.NORMAL;
        BaseColor? color = null;
        string? fontName = null;

        if (runPr != null)
        {
            if (runPr.FontSize != null && runPr.FontSize.Val.HasValue)
                fontSize = runPr.FontSize.Val.Value / 100f;

            if (runPr.Bold != null && (runPr.Bold.Val == null || runPr.Bold.Val.Value))
                fontStyle |= iTextFont.BOLD;

            if (runPr.Italic != null && (runPr.Italic.Val == null || runPr.Italic.Val.Value))
                fontStyle |= iTextFont.ITALIC;

            // color may be specified as rgb
            var rgb = runPr.Descendants<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
            if (!string.IsNullOrEmpty(rgb))
            {
                try
                {
                    color = new BaseColor(int.Parse(rgb, System.Globalization.NumberStyles.HexNumber));
                }
                catch { }
            }

            // latin font face
            var latin = runPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.LatinFont>()?.Typeface?.Value;
            if (!string.IsNullOrEmpty(latin))
                fontName = latin;
        }

        if (fontName != null && _fontNameMap.TryGetValue(fontName, out var mapped))
            fontName = mapped;

        if (fontName == null)
            fontName = GetDefaultFontName();

        return FontFactory.GetFont(fontName, fontSize, fontStyle, color ?? BaseColor.Black);
    }

    private static iTextFont? TryGetTrueBoldFont(string fontName, float fontSize, int fontStyle, BaseColor? color)
    {
        try
        {
            static bool LooksBoldName(string name) =>
                name.Contains("Bold", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Black", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Heavy", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Semibold", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Semi Bold", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Demibold", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("Demi Bold", StringComparison.OrdinalIgnoreCase);

            var candidates = new System.Collections.Generic.List<string>
            {
                fontName + " Bold",
                fontName + "-Bold",
                fontName + "Bold",
                fontName + " Semibold",
                fontName + "-Semibold",
                fontName + " DemiBold",
                fontName + "-DemiBold"
            };

            foreach (var c in candidates)
            {
                if (!FontFactory.IsRegistered(c)) continue;
                var f = FontFactory.GetFont(c, fontSize, fontStyle, color ?? BaseColor.Black);
                if (f.BaseFont != null && FontLooksBold(f))
                    return NormalizeBoldStyleIfNeeded(f, true, fontSize, fontStyle, color);
            }

            foreach (var regName in FontFactory.RegisteredFonts)
            {
                if (!LooksBoldName(regName)) continue;
                if (!regName.Contains(fontName, StringComparison.OrdinalIgnoreCase)) continue;
                var f = FontFactory.GetFont(regName, fontSize, fontStyle, color ?? BaseColor.Black);
                if (f.BaseFont != null && FontLooksBold(f))
                    return NormalizeBoldStyleIfNeeded(f, true, fontSize, fontStyle, color);
            }

            return null;
        }
        catch
        {
            return null;
        }
    }

    private static iTextFont NormalizeBoldStyleIfNeeded(iTextFont font, bool wantsBold, float fontSize, int requestedStyle, BaseColor? color)
    {
        if (!wantsBold) return font;
        if (!FontLooksBold(font)) return font;
        var style = requestedStyle & ~iTextFont.BOLD;
        return new iTextFont(font.Family, fontSize, style, color ?? BaseColor.Black);
    }

    private static bool FontLooksBold(iTextFont? font)
    {
        if (font == null) return false;
        var family = font.Family;
        if (string.IsNullOrWhiteSpace(family)) return false;
        if (family.Contains("Bold", StringComparison.OrdinalIgnoreCase)) return true;
        if (family.Contains("Black", StringComparison.OrdinalIgnoreCase)) return true;
        if (family.Contains("Heavy", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    /// <summary>
    /// ??????????? Font(?????)
    /// </summary>
    public iTextFont GetFont(float size, int style = iTextFont.NORMAL, BaseColor? color = null)
    {
        var cacheKey = $"simple_{size}_{style}_{color?.ToArgb() ?? 0}";
        return _fontCache.GetOrAdd(cacheKey, _ =>
        {
            var fontName = GetDefaultFontName();
            return new iTextFont(fontName, size, style, color ?? BaseColor.Black);
        });
    }
}
