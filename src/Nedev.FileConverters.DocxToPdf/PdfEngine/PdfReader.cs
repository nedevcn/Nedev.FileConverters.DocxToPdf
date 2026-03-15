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
    }

    public void Dispose()
    {
        // ????
    }
}

/// <summary>
/// PDF编辑工具（用于添加水印等）
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
        // Local helper to write bytes to the output stream
        void WriteBytes(byte[] bytes) => _outputStream.Write(bytes, 0, bytes.Length);

        // write the entire original PDF bytes
        var originalBytes = _reader.GetRawBytes();
        _outputStream.Write(originalBytes, 0, originalBytes.Length);

        // track new objects and their offsets for incremental xref
        var newOffsets = new Dictionary<int, long>();
        int maxObj = _reader.ObjectOffsets.Keys.DefaultIfEmpty(0).Max();

        // helper to write a new object with given body text (dictionary or stream) and record offset
        void WriteNewObject(int objNum, string body)
        {
            newOffsets[objNum] = _outputStream.Position;
            var header = $"{objNum} 0 obj\n";
            WriteBytes(System.Text.Encoding.Latin1.GetBytes(header));
            WriteBytes(System.Text.Encoding.Latin1.GetBytes(body));
            WriteBytes(System.Text.Encoding.Latin1.GetBytes("\nendobj\n\n"));
        }

        // write stream object helper
        int WriteStream(int objNum, string content)
        {
            var bytes = System.Text.Encoding.Latin1.GetBytes(content);
            var dict = $"<< /Length {bytes.Length} >>\nstream\n";
            newOffsets[objNum] = _outputStream.Position;
            var header = $"{objNum} 0 obj\n" + dict;
            WriteBytes(System.Text.Encoding.Latin1.GetBytes(header));
            _outputStream.Write(bytes, 0, bytes.Length);
            WriteBytes(System.Text.Encoding.Latin1.GetBytes("\nendstream\nendobj\n\n"));
            return objNum;
        }

        // for each page prepare new streams and page object modifications
        for (int page = 1; page <= _overContent.Count; page++)
        {
            var over = _overContent[page - 1]?.GetContent();
            var under = _underContent[page - 1]?.GetContent();
            if (string.IsNullOrEmpty(over) && string.IsNullOrEmpty(under))
                continue;

            List<int> extraStreams = new();
            if (!string.IsNullOrEmpty(under))
            {
                maxObj++;
                WriteStream(maxObj, under);
                extraStreams.Add(maxObj);
            }
            if (!string.IsNullOrEmpty(over))
            {
                maxObj++;
                WriteStream(maxObj, over);
                extraStreams.Add(maxObj);
            }

            // modify corresponding page object by writing a new version
            int pageObj = _reader.GetPageObjectNumber(page);
            if (pageObj > 0)
            {
                var orig = _reader.GetObjectText(pageObj) ?? "";
                // insert extras into /Contents
                string modified = orig;
                int ci = modified.IndexOf("/Contents");
                if (ci >= 0)
                {
                    int start = ci + "/Contents".Length;
                    while (start < modified.Length && char.IsWhiteSpace(modified[start])) start++;
                    if (start < modified.Length && modified[start] == '[')
                    {
                        int end = modified.IndexOf(']', start);
                        if (end > start)
                        {
                            string inside = modified.Substring(start + 1, end - start - 1);
                            foreach (var s in extraStreams)
                                inside += " " + s + " 0 R";
                            modified = modified.Substring(0, start + 1) + inside + modified.Substring(end);
                        }
                    }
                    else
                    {
                        // single reference
                        int end = modified.IndexOf("R", start);
                        if (end > start)
                        {
                            string existing = modified.Substring(start, end - start + 1);
                            string arr = "[ " + existing;
                            foreach (var s in extraStreams)
                                arr += " " + s + " 0 R";
                            arr += " ]";
                            modified = modified.Substring(0, start) + arr + modified.Substring(end + 1);
                        }
                    }
                }
                else
                {
                    int pos = modified.LastIndexOf(">>");
                    if (pos >= 0)
                    {
                        modified = modified.Substring(0, pos) + " /Contents ";
                        if (extraStreams.Count == 1)
                            modified += extraStreams[0] + " 0 R ";
                        else
                        {
                            modified += "[";
                            foreach (var s in extraStreams) modified += " " + s + " 0 R";
                            modified += " ] ";
                        }
                        modified += modified.Substring(pos);
                    }
                }
                WriteNewObject(pageObj, modified);
            }
        }

        // write incremental xref table
        long xrefStart = _outputStream.Position;
        var allNewObjs = newOffsets.Keys.OrderBy(n => n).ToList();
        var headerLine = $"xref\n0 {allNewObjs.Max() + 1}\n";
        WriteBytes(System.Text.Encoding.Latin1.GetBytes(headerLine));
        // zero entry
        WriteBytes(System.Text.Encoding.Latin1.GetBytes("0000000000 65535 f \r\n"));
        for (int i = 1; i <= allNewObjs.Max(); i++)
        {
            if (newOffsets.TryGetValue(i, out var off))
            {
                var entry = $"{off:D10} 00000 n \r\n";
                WriteBytes(System.Text.Encoding.Latin1.GetBytes(entry));
            }
            else
            {
                WriteBytes(System.Text.Encoding.Latin1.GetBytes("0000000000 65535 f \r\n"));
            }
        }

        // write trailer including Prev if available
        var trailerSb = new System.Text.StringBuilder();
        trailerSb.Append("trailer\n<< ");
        if (_reader.XrefOffset.HasValue)
            trailerSb.Append($"/Prev {_reader.XrefOffset.Value} ");
        trailerSb.Append($"/Size {allNewObjs.Max() + 1} >>\n");
        trailerSb.Append($"startxref\n{xrefStart}\n%%EOF\n");
        WriteBytes(System.Text.Encoding.Latin1.GetBytes(trailerSb.ToString()));

        _outputStream.Flush();
    }

    public void Dispose()
    {
        Close();
    }
}
