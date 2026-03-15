namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// PDF???
/// </summary>
public class BaseColor
{
    public byte R { get; }
    public byte G { get; }
    public byte B { get; }
    public byte A { get; }

    public static readonly BaseColor Black = new(0, 0, 0);
    public static readonly BaseColor White = new(255, 255, 255);
    public static readonly BaseColor Red = new(255, 0, 0);
    public static readonly BaseColor Green = new(0, 255, 0);
    public static readonly BaseColor Blue = new(0, 0, 255);
    public static readonly BaseColor Yellow = new(255, 255, 0);
    public static readonly BaseColor Cyan = new(0, 255, 255);
    public static readonly BaseColor Magenta = new(255, 0, 255);
    public static readonly BaseColor Gray = new(128, 128, 128);
    public static readonly BaseColor LightGray = new(211, 211, 211);
    public static readonly BaseColor DarkGray = new(169, 169, 169);

    public BaseColor(byte r, byte g, byte b, byte a = 255)
    {
        R = r;
        G = g;
        B = b;
        A = a;
    }

    public BaseColor(int r, int g, int b, int a = 255)
    {
        R = (byte)Math.Clamp(r, 0, 255);
        G = (byte)Math.Clamp(g, 0, 255);
        B = (byte)Math.Clamp(b, 0, 255);
        A = (byte)Math.Clamp(a, 0, 255);
    }

    public BaseColor(float r, float g, float b, float a = 1.0f)
    {
        R = (byte)Math.Clamp(Math.Round(r * 255), 0, 255);
        G = (byte)Math.Clamp(Math.Round(g * 255), 0, 255);
        B = (byte)Math.Clamp(Math.Round(b * 255), 0, 255);
        A = (byte)Math.Clamp(Math.Round(a * 255), 0, 255);
    }

    /// <summary>
    /// 从 ARGB 整数创建颜色 (0xAARRGGBB 格式)
    /// </summary>
    public BaseColor(int argb)
    {
        A = (byte)((argb >> 24) & 0xFF);
        R = (byte)((argb >> 16) & 0xFF);
        G = (byte)((argb >> 8) & 0xFF);
        B = (byte)(argb & 0xFF);
    }

    public int ToArgb() => (A << 24) | (R << 16) | (G << 8) | B;

    public float[] ToRgbFloats() => [R / 255f, G / 255f, B / 255f];

    public string ToHex() => $"{R:X2}{G:X2}{B:X2}";

    public static BaseColor FromHex(string hex)
    {
        if (string.IsNullOrWhiteSpace(hex)) return Black;
        hex = hex.TrimStart('#');
        if (hex.Length == 6 &&
            int.TryParse(hex[..2], System.Globalization.NumberStyles.HexNumber, null, out var r) &&
            int.TryParse(hex[2..4], System.Globalization.NumberStyles.HexNumber, null, out var g) &&
            int.TryParse(hex[4..6], System.Globalization.NumberStyles.HexNumber, null, out var b))
        {
            return new BaseColor(r, g, b);
        }
        return Black;
    }
}
