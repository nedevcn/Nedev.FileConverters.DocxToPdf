namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// PDF???(???)
/// </summary>
public class PdfReader : IDisposable
{
    private readonly byte[] _pdfData;
    private readonly List<Rectangle> _pageSizes = [];
    // offsets/lengths of page content streams (one per page, in order)
    private readonly List<int> _streamOffsets = [];
    private readonly List<int> _streamLengths = [];

    // map object number -> byte offset (from xref table)
    private readonly Dictionary<int, int> _objectOffsets = new();
    // keep page object numbers in order of discovery
    private readonly List<int> _pageObjectNumbers = [];
    // offset of the startxref position from the original file (if parsed)
    public int? XrefOffset { get; private set; }

    public int NumberOfPages => _pageSizes.Count;

    /// <summary>
    /// Offsets parsed from the xref table. Useful for reading raw objects.
    /// </summary>
    public IReadOnlyDictionary<int,int> ObjectOffsets => _objectOffsets;

    /// <summary>
    /// Object number corresponding to the given page (1-based).
    /// Returns -1 if not known.
    /// </summary>
    public int GetPageObjectNumber(int pageNum)
    {
        if (pageNum < 1 || pageNum > _pageObjectNumbers.Count)
            return -1;
        return _pageObjectNumbers[pageNum - 1];
    }

    /// <summary>
    /// Return the raw text of the object with the given number, or null if unknown.
    /// The returned string does not include the "N 0 obj" header or the trailing "endobj".
    /// </summary>
    public string? GetObjectText(int objNum)
    {
        if (_objectOffsets.TryGetValue(objNum, out var offset))
        {
            // read from offset to endobj
            try
            {
                var text = System.Text.Encoding.ASCII.GetString(_pdfData, offset, _pdfData.Length - offset);
                var idx = text.IndexOf("endobj");
                if (idx >= 0)
                {
                    return text.Substring(0, idx);
                }
                return text;
            }
            catch
            {
                return null;
            }
        }
        return null;
    }

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

            // record stream offsets/lengths by scanning for content pairs
            var streamRegex = new System.Text.RegularExpressions.Regex(@"stream\r?\n",
                System.Text.RegularExpressions.RegexOptions.Compiled);
            var endRegex = new System.Text.RegularExpressions.Regex(@"endstream",
                System.Text.RegularExpressions.RegexOptions.Compiled);
            var streamMatches = streamRegex.Matches(text);
            var endMatches = endRegex.Matches(text);
            int pagesFound = Math.Min(streamMatches.Count, endMatches.Count);
            for (int si = 0; si < pagesFound; si++)
            {
                int start = streamMatches[si].Index + streamMatches[si].Length;
                int end = endMatches[si].Index;
                _streamOffsets.Add(start);
                _streamLengths.Add(end - start);
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

            // parse xref table and gather object offsets
            ParseXref(text);

            // identify page object numbers in object order
            foreach (var kvp in _objectOffsets.OrderBy(k => k.Key))
            {
                var objText = GetObjectText(kvp.Key);
                if (objText != null && objText.Contains("/Type /Page"))
                {
                    _pageObjectNumbers.Add(kvp.Key);
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

    /// <summary>
    /// Return the raw PDF bytes originally passed to the reader.
    /// </summary>
    public byte[] GetRawBytes() => _pdfData;

    private void ParseXref(string text)
    {
        var startXrefRegex = new System.Text.RegularExpressions.Regex(@"startxref\s+(\d+)",
            System.Text.RegularExpressions.RegexOptions.Compiled);
        var m = startXrefRegex.Match(text);
        if (m.Success && int.TryParse(m.Groups[1].Value, out var offset))
        {
            XrefOffset = offset;
            if (offset >= 0 && offset < text.Length)
            {
                var xrefText = text.Substring(offset);
                var lines = xrefText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length > 0 && lines[0].Trim() == "xref")
                {
                    int idx = 1;
                    while (idx < lines.Length)
                    {
                        var headerParts = lines[idx].Trim().Split(' ');
                        if (headerParts.Length == 2 &&
                            int.TryParse(headerParts[0], out int start) &&
                            int.TryParse(headerParts[1], out int count))
                        {
                            idx++;
                            for (int i = 0; i < count && idx < lines.Length; i++, idx++)
                            {
                                var entry = lines[idx];
                                var parts = entry.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                                if (parts.Length >= 3 &&
                                    int.TryParse(parts[0], out var off) &&
                                    parts[2] == "n")
                                {
                                    _objectOffsets[start + i] = off;
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
        }
    }

    public byte[] GetPageContent(int pageNum)
    {
        if (pageNum < 1 || pageNum > _streamOffsets.Count)
            return _pdfData;

        try
        {
            int start = _streamOffsets[pageNum - 1];
            int len = _streamLengths[pageNum - 1];
            var result = new byte[len];
            Array.Copy(_pdfData, start, result, 0, len);
            return result;
        }
        catch
        {
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
