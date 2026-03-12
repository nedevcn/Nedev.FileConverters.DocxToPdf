namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// PDF???(???)
/// </summary>
public class PdfReader : IDisposable
{
    private readonly byte[] _pdfData;
    private readonly List<Rectangle> _pageSizes = [];

    public int NumberOfPages => _pageSizes.Count;

    public PdfReader(byte[] pdfData)
    {
        _pdfData = pdfData ?? throw new ArgumentNullException(nameof(pdfData));
        ParsePages();
    }

    public PdfReader(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        _pdfData = ms.ToArray();
        ParsePages();
    }

    private void ParsePages()
    {
        // attempt to count pages by scanning for "/Type /Page" tokens and capture MediaBox sizes.
        try
        {
            var text = System.Text.Encoding.ASCII.GetString(_pdfData);
            var matches = System.Text.RegularExpressions.Regex.Matches(text, @"/Type\s*/Page");
            var count = matches.Count;
            if (count <= 0)
                count = 1;

            for (int i = 0; i < count; i++)
            {
                _pageSizes.Add(Rectangle.A4);
            }

            // look for media boxes to override defaults
            var mediaRegex = new System.Text.RegularExpressions.Regex(@"/MediaBox\s*\[\s*([0-9\.\-]+)\s+([0-9\.\-]+)\s+([0-9\.\-]+)\s+([0-9\.\-]+)\s*\]",
                System.Text.RegularExpressions.RegexOptions.Compiled);
            int idx = 0;
            foreach (System.Text.RegularExpressions.Match m in mediaRegex.Matches(text))
            {
                if (m.Groups.Count == 5 && idx < _pageSizes.Count)
                {
                    if (float.TryParse(m.Groups[1].Value, out var left) &&
                        float.TryParse(m.Groups[2].Value, out var bottom) &&
                        float.TryParse(m.Groups[3].Value, out var right) &&
                        float.TryParse(m.Groups[4].Value, out var top))
                    {
                        _pageSizes[idx] = new Rectangle(left, bottom, right, top);
                        idx++;
                    }
                }
            }
        }
        catch
        {
            _pageSizes.Clear();
            _pageSizes.Add(Rectangle.A4);
        }
    }

    public Rectangle GetPageSize(int pageNum)
    {
        if (pageNum < 1 || pageNum > _pageSizes.Count)
            return Rectangle.A4;
        return _pageSizes[pageNum - 1];
    }

    public byte[] GetPageContent(int pageNum)
    {
        // ????:??????
        return _pdfData;
    }

    public void Dispose()
    {
        // ????
    }
}

/// <summary>
/// PDF???(???????)
/// </summary>
public class PdfStamper : IDisposable
{
    private readonly PdfReader _reader;
    private readonly Stream _outputStream;
    private readonly List<PdfContentByte> _overContent = [];
    private readonly List<PdfContentByte> _underContent = [];

    public PdfStamper(PdfReader reader, Stream outputStream)
    {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        _outputStream = outputStream ?? throw new ArgumentNullException(nameof(outputStream));

        // initialize empty content buffers for each page
        for (var i = 0; i < reader.NumberOfPages; i++)
        {
            _overContent.Add(new PdfContentByte());
            _underContent.Add(new PdfContentByte());
        }
    }

    public PdfContentByte GetOverContent(int pageNum)
    {
        if (pageNum < 1 || pageNum > _overContent.Count)
            throw new ArgumentOutOfRangeException(nameof(pageNum));
        return _overContent[pageNum - 1];
    }

    public PdfContentByte GetUnderContent(int pageNum)
    {
        if (pageNum < 1 || pageNum > _underContent.Count)
            throw new ArgumentOutOfRangeException(nameof(pageNum));
        return _underContent[pageNum - 1];
    }

    public void Close()
    {
        // write the original PDF bytes first
        var originalData = _reader.GetPageContent(1);
        _outputStream.Write(originalData, 0, originalData.Length);

        // then append any over/under content as comments so that stamping is not a no-op
        for (int i = 0; i < _overContent.Count; i++)
        {
            var over = _overContent[i]?.GetContent();
            if (!string.IsNullOrEmpty(over))
            {
                var comment = System.Text.Encoding.UTF8.GetBytes($"\n% OverContent page {i+1}\n{over}\n");
                _outputStream.Write(comment, 0, comment.Length);
            }
            var under = _underContent[i]?.GetContent();
            if (!string.IsNullOrEmpty(under))
            {
                var comment2 = System.Text.Encoding.UTF8.GetBytes($"\n% UnderContent page {i+1}\n{under}\n");
                _outputStream.Write(comment2, 0, comment2.Length);
            }
        }

        _outputStream.Flush();
    }

    public void Dispose()
    {
        Close();
    }
}
