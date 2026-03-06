using SkiaSharp;

namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// PDF???
/// </summary>
public class Image : IElement
{
    public const int UNDERLYING = 1;
    public const int ALIGN_CENTER = 1;
    public const int ALIGN_LEFT = 0;
    public const int ALIGN_RIGHT = 2;

    private byte[] _imageData;
    private float _scaledWidth;
    private float _scaledHeight;
    private float _absoluteX = -1;
    private float _absoluteY = -1;

    public int Type => 10;
    public bool IsContent() => true;
    public bool IsNestable() => false;

    public int Alignment { get; set; } = ALIGN_CENTER;
    public float OriginalWidth { get; private set; }
    public float OriginalHeight { get; private set; }

    public float ScaledWidth => _scaledWidth;
    public float ScaledHeight => _scaledHeight;

    public float AbsoluteX => _absoluteX;
    public float AbsoluteY => _absoluteY;

    public bool HasAbsolutePosition => _absoluteX >= 0 && _absoluteY >= 0;

    public byte[] ImageData => _imageData;
    public SKEncodedImageFormat Format { get; private set; }

    private Image(byte[] data, float width, float height, SKEncodedImageFormat format)
    {
        _imageData = data;
        OriginalWidth = width;
        OriginalHeight = height;
        _scaledWidth = width;
        _scaledHeight = height;
        Format = format;
    }

    public static Image? GetInstance(byte[] imageData)
    {
        try
        {
            using var stream = new SKMemoryStream(imageData);
            using var codec = SKCodec.Create(stream);
            if (codec == null) return null;

            var info = codec.Info;
            var format = codec.EncodedFormat;

            return new Image(imageData, info.Width, info.Height, format);
        }
        catch
        {
            return null;
        }
    }

    public static Image? GetInstance(string filePath)
    {
        if (!File.Exists(filePath)) return null;
        try
        {
            var data = File.ReadAllBytes(filePath);
            return GetInstance(data);
        }
        catch
        {
            return null;
        }
    }

    public void ScaleAbsolute(float width, float height)
    {
        _scaledWidth = width;
        _scaledHeight = height;
    }

    public void ScaleToFit(float maxWidth, float maxHeight)
    {
        if (OriginalWidth <= 0 || OriginalHeight <= 0)
        {
            _scaledWidth = maxWidth;
            _scaledHeight = maxHeight;
            return;
        }
        var widthRatio = maxWidth / OriginalWidth;
        var heightRatio = maxHeight / OriginalHeight;
        var ratio = Math.Min(widthRatio, heightRatio);

        _scaledWidth = OriginalWidth * ratio;
        _scaledHeight = OriginalHeight * ratio;
    }

    public void SetAbsolutePosition(float x, float y)
    {
        _absoluteX = x;
        _absoluteY = y;
    }

    public byte[] GetPngData()
    {
        if (Format == SKEncodedImageFormat.Png) return _imageData;

        // ???PNG
        try
        {
            using var stream = new SKMemoryStream(_imageData);
            using var bitmap = SKBitmap.Decode(stream);
            if (bitmap == null) return _imageData;

            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            return data.ToArray();
        }
        catch
        {
            return _imageData;
        }
    }

    public byte[] GetJpegData(int quality = 90)
    {
        if (Format == SKEncodedImageFormat.Jpeg) return _imageData;

        try
        {
            using var stream = new SKMemoryStream(_imageData);
            using var codec = SKCodec.Create(stream);
            if (codec == null) return _imageData;

            // ?????????,??????????????,???JPEG?????/???
            var info = new SKImageInfo(codec.Info.Width, codec.Info.Height, SKColorType.Bgra8888, SKAlphaType.Premul);
            using var bitmap = new SKBitmap(info);
            
            using (var canvas = new SKCanvas(bitmap))
            {
                canvas.Clear(SKColors.White);
                using var originalBitmap = SKBitmap.Decode(codec);
                if (originalBitmap != null)
                {
                    canvas.DrawBitmap(originalBitmap, 0, 0);
                }
            }

            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Jpeg, quality);
            return data.ToArray();
        }
        catch
        {
            return _imageData;
        }
    }

    /// <summary>
    /// ?????Base64??(??????)
    /// </summary>
    public string GetBase64Data()
    {
        return Convert.ToBase64String(_imageData);
    }

    /// <summary>
    /// ????????
    /// </summary>
    public string GetFormatName()
    {
        return Format switch
        {
            SKEncodedImageFormat.Png => "PNG",
            SKEncodedImageFormat.Jpeg => "JPEG",
            SKEncodedImageFormat.Gif => "GIF",
            SKEncodedImageFormat.Bmp => "BMP",
            SKEncodedImageFormat.Webp => "WEBP",
            _ => "UNKNOWN"
        };
    }
}
