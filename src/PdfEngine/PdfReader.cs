namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// PDF读取器（简化版）
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
        // 简化实现：假设所有页面都是A4
        // 实际应该解析PDF结构
        _pageSizes.Add(Rectangle.A4);
    }

    public Rectangle GetPageSize(int pageNum)
    {
        if (pageNum < 1 || pageNum > _pageSizes.Count)
            return Rectangle.A4;
        return _pageSizes[pageNum - 1];
    }

    public byte[] GetPageContent(int pageNum)
    {
        // 简化实现：返回原始数据
        return _pdfData;
    }

    public void Dispose()
    {
        // 清理资源
    }
}

/// <summary>
/// PDF盖章器（用于添加水印等）
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

        // 初始化内容层
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
        // 合并原始PDF和新增内容
        // 简化实现：直接复制原始数据
        var originalData = _reader.GetPageContent(1);
        _outputStream.Write(originalData, 0, originalData.Length);
        _outputStream.Flush();
    }

    public void Dispose()
    {
        Close();
    }
}
