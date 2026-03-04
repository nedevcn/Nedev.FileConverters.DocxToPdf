using System.Collections.Concurrent;

namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// 兼容性BaseFont
/// </summary>
public class BaseFont
{
    public const string IDENTITY_H = "Identity-H";
    public const string CP1252 = "Cp1252";
    public const bool EMBEDDED = true;
    public const bool NOT_EMBEDDED = false;

    public const int ASCENT = 1;
    public const int DESCENT = 2;
    public const int CAPHEIGHT = 3;

    public string PostscriptFontName { get; }

    public BaseFont(string name)
    {
        PostscriptFontName = name;
    }

    public static BaseFont CreateFont(string name, string encoding, bool embedded)
    {
        return new BaseFont(name);
    }

    public float GetFontDescriptor(int key, float fontSize)
    {
        return key switch
        {
            ASCENT => fontSize * 0.8f,
            DESCENT => -fontSize * 0.2f,
            CAPHEIGHT => fontSize * 0.7f,
            _ => fontSize
        };
    }

    public float GetWidthPoint(string text, float fontSize)
    {
        // 改进的宽度计算，区分中英文
        float width = 0;
        foreach (var c in text)
        {
            if (c >= '\u4e00' && c <= '\u9fff')
            {
                width += fontSize;
            }
            else if (c >= '\u3000' && c <= '\u303f')
            {
                width += fontSize;
            }
            else if (c >= '\uff00' && c <= '\uffef')
            {
                width += fontSize;
            }
            else
            {
                width += fontSize * 0.5f;
            }
        }
        return width;
    }
}

/// <summary>
/// PDF字体类
/// </summary>
public class Font
{
    public const int NORMAL = 0;
    public const int BOLD = 1;
    public const int ITALIC = 2;
    public const int UNDERLINE = 4;
    public const int STRIKETHRU = 8;

    public string Family { get; }
    public float Size { get; }
    public int Style { get; }
    public BaseColor Color { get; }

    private static readonly ConcurrentDictionary<string, Font> _fontCache = new();

    public Font(string family, float size, int style = NORMAL, BaseColor? color = null)
    {
        Family = family ?? "Helvetica";
        Size = size > 0 ? size : 12;
        Style = style;
        Color = color ?? BaseColor.Black;
    }

    public bool IsBold => (Style & BOLD) != 0;
    public bool IsItalic => (Style & ITALIC) != 0;
    public bool IsUnderline => (Style & UNDERLINE) != 0;
    public bool IsStrikethru => (Style & STRIKETHRU) != 0;

    public Font WithSize(float size) => new(Family, size, Style, Color);
    public Font WithStyle(int style) => new(Family, Size, style, Color);
    public Font WithColor(BaseColor color) => new(Family, Size, Style, color);

    public float GetWidthPoint(string text)
    {
        // 改进的宽度计算：
        // - 中文字符（\u4e00-\u9fff）：全宽，约等于字体大小
        // - 其他字符（英文、数字等）：半宽，约等于字体大小的一半
        float width = 0;
        foreach (var c in text)
        {
            if (c >= '\u4e00' && c <= '\u9fff')
            {
                // CJK统一汉字
                width += Size;
            }
            else if (c >= '\u3000' && c <= '\u303f')
            {
                // CJK标点符号
                width += Size;
            }
            else if (c >= '\uff00' && c <= '\uffef')
            {
                // 全角ASCII/半角片假名
                width += Size;
            }
            else
            {
                // 其他字符（英文、数字等）
                width += Size * 0.5f;
            }
        }
        return width;
    }

    // 兼容性属性：BaseFont
    public BaseFont BaseFont => new(Family);

    public static Font GetFont(string family, float size, int style = NORMAL, BaseColor? color = null)
    {
        var key = $"{family}_{size}_{style}_{color?.ToArgb() ?? 0}";
        return _fontCache.GetOrAdd(key, _ => new Font(family, size, style, color));
    }

    public static Font Helvetica(float size, int style = NORMAL, BaseColor? color = null) =>
        GetFont("Helvetica", size, style, color);

    public static Font Times(float size, int style = NORMAL, BaseColor? color = null) =>
        GetFont("Times-Roman", size, style, color);

    public static Font Courier(float size, int style = NORMAL, BaseColor? color = null) =>
        GetFont("Courier", size, style, color);
}

/// <summary>
/// 字体工厂
/// </summary>
public static class FontFactory
{
    private static readonly HashSet<string> _registeredFonts = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, string> _fontPathMap = new();

    static FontFactory()
    {
        // 注册标准字体
        _registeredFonts.Add("Helvetica");
        _registeredFonts.Add("Helvetica-Bold");
        _registeredFonts.Add("Helvetica-Oblique");
        _registeredFonts.Add("Helvetica-BoldOblique");
        _registeredFonts.Add("Times-Roman");
        _registeredFonts.Add("Times-Bold");
        _registeredFonts.Add("Times-Italic");
        _registeredFonts.Add("Times-BoldItalic");
        _registeredFonts.Add("Courier");
        _registeredFonts.Add("Courier-Bold");
        _registeredFonts.Add("Courier-Oblique");
        _registeredFonts.Add("Courier-BoldOblique");
        _registeredFonts.Add("Symbol");
        _registeredFonts.Add("ZapfDingbats");
    }

    public static void RegisterDirectory(string directory)
    {
        if (!Directory.Exists(directory)) return;

        foreach (var file in Directory.GetFiles(directory, "*.ttf"))
        {
            Register(file);
        }
        foreach (var file in Directory.GetFiles(directory, "*.ttc"))
        {
            Register(file);
        }
        foreach (var file in Directory.GetFiles(directory, "*.otf"))
        {
            Register(file);
        }
    }

    public static void Register(string fontPath)
    {
        if (!File.Exists(fontPath)) return;

        var fontName = Path.GetFileNameWithoutExtension(fontPath);
        _registeredFonts.Add(fontName);
        _fontPathMap[fontName] = fontPath;
    }

    public static bool IsRegistered(string fontName)
    {
        return _registeredFonts.Contains(fontName);
    }

    public static Font GetFont(string family, float size, int style = Font.NORMAL, BaseColor? color = null)
    {
        return Font.GetFont(family, size, style, color);
    }

    public static IEnumerable<string> RegisteredFonts => _registeredFonts;
}
