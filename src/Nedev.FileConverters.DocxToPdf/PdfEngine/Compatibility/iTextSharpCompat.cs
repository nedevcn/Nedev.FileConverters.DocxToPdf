// iTextSharp????
// ???iTextSharp???API,?????PDF????

using Nedev.FileConverters.DocxToPdf.PdfEngine;

// ????,????
using iTextDocument = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfDocument;
using iTextWriter = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfWriter;
using iTextRectangle = Nedev.FileConverters.DocxToPdf.PdfEngine.Rectangle;
using iTextBaseColor = Nedev.FileConverters.DocxToPdf.PdfEngine.BaseColor;
using iTextFont = Nedev.FileConverters.DocxToPdf.PdfEngine.Font;
using iTextChunk = Nedev.FileConverters.DocxToPdf.PdfEngine.Chunk;
using iTextParagraph = Nedev.FileConverters.DocxToPdf.PdfEngine.Paragraph;
using iTextPhrase = Nedev.FileConverters.DocxToPdf.PdfEngine.Phrase;
using iTextImage = Nedev.FileConverters.DocxToPdf.PdfEngine.Image;
using iTextPdfPTable = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfPTable;
using iTextPdfPCell = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfPCell;
using iTextList = Nedev.FileConverters.DocxToPdf.PdfEngine.List;
using iTextListItem = Nedev.FileConverters.DocxToPdf.PdfEngine.ListItem;
using iTextElement = Nedev.FileConverters.DocxToPdf.PdfEngine.IElement;
using iTextPdfContentByte = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfContentByte;
using iTextColumnText = Nedev.FileConverters.DocxToPdf.PdfEngine.ColumnText;
using iTextPdfOutline = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfOutline;
using iTextPdfDestination = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfDestination;
using iTextPdfPageEventHelper = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfPageEventHelper;
using iTextPdfReader = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfReader;
using iTextPdfStamper = Nedev.FileConverters.DocxToPdf.PdfEngine.PdfStamper;
using iTextFloatingObject = Nedev.FileConverters.DocxToPdf.PdfEngine.FloatingObject;
using iTextWrappingStyle = Nedev.FileConverters.DocxToPdf.PdfEngine.WrappingStyle;

namespace Nedev.FileConverters.DocxToPdf.PdfEngine.Compatibility;

/// <summary>
/// ???????
/// </summary>
public static class iTextSharpCompatExtensions
{
    public static iTextRectangle Rotate(this iTextRectangle rect) => rect.Rotate();
}

/// <summary>
/// ??????
/// </summary>
public static class PageSize
{
    public static iTextRectangle A4 => iTextRectangle.A4;
    public static iTextRectangle A3 => iTextRectangle.A3;
    public static iTextRectangle A5 => iTextRectangle.A5;
    public static iTextRectangle LETTER => iTextRectangle.Letter;
    public static iTextRectangle LEGAL => iTextRectangle.Legal;
}

/// <summary>
/// ???????
/// </summary>
public static class iTextFontFactory
{
    public const string HELVETICA = "Helvetica";
    public const string TIMES_ROMAN = "Times-Roman";
    public const string COURIER = "Courier";
    public const string SYMBOL = "Symbol";
    public const string ZAPFDINGBATS = "ZapfDingbats";

    public static void RegisterDirectory(string directory)
    {
        FontFactory.RegisterDirectory(directory);
    }

    public static void Register(string fontPath)
    {
        FontFactory.Register(fontPath);
    }

    public static bool IsRegistered(string fontName)
    {
        return FontFactory.IsRegistered(fontName);
    }

    public static iTextFont GetFont(string family, float size, int style = iTextFont.NORMAL, iTextBaseColor? color = null)
    {
        return FontFactory.GetFont(family, size, style, color);
    }

    public static iTextFont GetFont(string family, string encoding, bool embedded, float size, int style, iTextBaseColor color)
    {
        return new iTextFont(family, size, style, color);
    }
}

/// <summary>
/// ???BaseFont
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
        return text.Length * fontSize * 0.5f;
    }
}

/// <summary>
/// ?????
/// </summary>
public static class PdfContentByteConstants
{
    public const int TEXT_RENDER_MODE_FILL = 0;
    public const int TEXT_RENDER_MODE_FILL_STROKE = 2;
}

/// <summary>
/// ???????(???)
/// </summary>
public interface IElementListener
{
    bool Add(IElement element);
}

/// <summary>
/// PDF??????(???)
/// </summary>
public interface IPdfPTableEvent
{
    void TableLayout(PdfPTable table, float[][] widths, float[] heights, int headerRows, int rowStart, PdfContentByte[] canvases);
}

/// <summary>
/// PDF??????
/// </summary>
public class PdfContentByteExt : iTextPdfContentByte
{
    public const int LINECANVAS = 0;

    public PdfContentByteExt() : base() { }
    public PdfContentByteExt(iTextWriter writer) : base() { }

    public void SetTextRenderMode(int mode, float strokeWidth, iTextBaseColor strokeColor)
    {
        // ????
    }

    public new void AddImage(iTextImage image)
    {
        DrawImage(image, image.AbsoluteX, image.AbsoluteY);
    }

    public new void AddImage(iTextImage image, float x, float y)
    {
        DrawImage(image, x, y);
    }
}
