using System.Text;

namespace Nedev.DocxToPdf.PdfEngine;

/// <summary>
/// TrueType字体解析器 - 解析字体文件并提取必要信息
/// </summary>
public class TrueTypeFont
{
    private byte[] _data;
    private int _pos;
    private bool _isTTC = false;
    private uint _ttcOffset = 0;

    public string FontName { get; private set; } = "SimSun";
    public int UnitsPerEm { get; private set; } = 1000;
    public short Ascent { get; private set; } = 800;
    public short Descent { get; private set; } = -200;
    public short CapHeight { get; private set; } = 700;
    public short StemV { get; private set; } = 80;
    public int Flags { get; private set; } = 32;
    public short ItalicAngle { get; private set; } = 0;
    public short XMin { get; private set; } = -100;
    public short YMin { get; private set; } = -200;
    public short XMax { get; private set; } = 1000;
    public short YMax { get; private set; } = 800;

    // 字形索引到Unicode的映射
    public Dictionary<int, char> GlyphToUnicode { get; private set; } = new();
    public Dictionary<char, int> UnicodeToGlyph { get; private set; } = new();

    public int[]? GlyphWidths { get; private set; }
    private int _numberOfHMetrics;

    public TrueTypeFont(byte[] data)
    {
        _data = data;
        Parse();
    }

    private void Parse()
    {
        _pos = 0;

        // 检查是否为TTC文件
        var tag = ReadTag();
        if (tag == "ttcf")
        {
            _isTTC = true;
            ParseTTCHeader();
        }
        else
        {
            // 普通TTF文件，重置位置
            _pos = 0;
        }

        // 读取字体头
        var sfntVersion = ReadUInt32();
        var numTables = ReadUInt16();
        var searchRange = ReadUInt16();
        var entrySelector = ReadUInt16();
        var rangeShift = ReadUInt16();

        // 读取表目录
        var tables = new Dictionary<string, (uint offset, uint length)>();
        for (int i = 0; i < numTables; i++)
        {
            var tableTag = ReadTag();
            var checksum = ReadUInt32();
            var offset = ReadUInt32();
            var length = ReadUInt32();
            tables[tableTag] = (offset + _ttcOffset, length);
        }

        // 解析head表
        if (tables.TryGetValue("head", out var headInfo))
        {
            ParseHeadTable(headInfo.offset);
        }

        // 提取 hmtx 表偏移
        uint hmtxOffset = 0;
        if (tables.TryGetValue("hmtx", out var hmtxInfo))
        {
            hmtxOffset = hmtxInfo.offset;
        }

        // 解析hhea表
        if (tables.TryGetValue("hhea", out var hheaInfo))
        {
            ParseHheaTable(hheaInfo.offset);
        }

        // 解析hmtx表
        if (hmtxOffset > 0 && _numberOfHMetrics > 0)
        {
            ParseHmtxTable(hmtxOffset);
        }

        // 解析name表获取字体名称
        if (tables.TryGetValue("name", out var nameInfo))
        {
            ParseNameTable(nameInfo.offset);
        }

        // 解析cmap表获取字符映射
        if (tables.TryGetValue("cmap", out var cmapInfo))
        {
            ParseCmapTable(cmapInfo.offset);
        }

        // 如果字体名称为空，使用默认值
        if (string.IsNullOrEmpty(FontName))
        {
            FontName = "SimSun";
        }
    }

    private void ParseTTCHeader()
    {
        // TTC Header
        var ttcVersion = ReadUInt32(); // Version
        var numFonts = ReadUInt32();   // Number of fonts in collection

        // 读取第一个字体的偏移量
        if (numFonts > 0)
        {
            _ttcOffset = ReadUInt32();
            _pos = (int)_ttcOffset;
        }
    }

    private void ParseHeadTable(uint offset)
    {
        _pos = (int)offset;
        var version = ReadUInt32();
        var fontRevision = ReadUInt32();
        var checkSumAdjustment = ReadUInt32();
        var magicNumber = ReadUInt32();
        var flags = ReadUInt16();
        UnitsPerEm = ReadUInt16();
        var created = ReadUInt64();
        var modified = ReadUInt64();
        XMin = ReadInt16();
        YMin = ReadInt16();
        XMax = ReadInt16();
        YMax = ReadInt16();
        var macStyle = ReadUInt16();
        var lowestRecPPEM = ReadUInt16();
        var fontDirectionHint = ReadInt16();
        var indexToLocFormat = ReadInt16();
        var glyphDataFormat = ReadInt16();
    }

    private void ParseHheaTable(uint offset)
    {
        _pos = (int)offset;
        var version = ReadUInt32();
        Ascent = ReadInt16();
        Descent = ReadInt16();
        var lineGap = ReadInt16();
        var advanceWidthMax = ReadUInt16();
        var minLeftSideBearing = ReadInt16();
        var minRightSideBearing = ReadInt16();
        var xMaxExtent = ReadInt16();
        var caretSlopeRise = ReadInt16();
        var caretSlopeRun = ReadInt16();
        var caretOffset = ReadInt16();
        var reserved1 = ReadInt16();
        var reserved2 = ReadInt16();
        var reserved3 = ReadInt16();
        var reserved4 = ReadInt16();
        var metricDataFormat = ReadInt16();
        _numberOfHMetrics = ReadUInt16();
    }

    private void ParseHmtxTable(uint offset)
    {
        _pos = (int)offset;
        GlyphWidths = new int[_numberOfHMetrics];
        for (int i = 0; i < _numberOfHMetrics; i++)
        {
            GlyphWidths[i] = ReadUInt16();
            ReadInt16(); // Left Side Bearing
        }
    }

    private void ParseNameTable(uint offset)
    {
        _pos = (int)offset;
        var format = ReadUInt16();
        var count = ReadUInt16();
        var stringOffset = ReadUInt16();

        for (int i = 0; i < count; i++)
        {
            var platformID = ReadUInt16();
            var encodingID = ReadUInt16();
            var languageID = ReadUInt16();
            var nameID = ReadUInt16();
            var length = ReadUInt16();
            var nameOffset = ReadUInt16();

            // 名称ID 4是完整字体名称
            if (nameID == 4 && string.IsNullOrEmpty(FontName))
            {
                var savePos = _pos;
                _pos = (int)offset + stringOffset + nameOffset;
                var nameBytes = ReadBytes(length);

                try
                {
                    if (platformID == 3) // Windows平台
                    {
                        if (encodingID == 1 || encodingID == 10) // Unicode
                        {
                            FontName = Encoding.BigEndianUnicode.GetString(nameBytes);
                        }
                    }
                    else if (platformID == 1) // Mac平台
                    {
                        FontName = Encoding.ASCII.GetString(nameBytes);
                    }
                }
                catch
                {
                    // 解析失败时使用默认值
                }

                _pos = savePos;
            }
        }
    }

    private void ParseCmapTable(uint offset)
    {
        _pos = (int)offset;
        var version = ReadUInt16();
        var numTables = ReadUInt16();

        for (int i = 0; i < numTables; i++)
        {
            var platformID = ReadUInt16();
            var encodingID = ReadUInt16();
            var subtableOffset = ReadUInt32();

            var savePos = _pos;
            ParseCmapSubtable(offset + subtableOffset, platformID, encodingID);
            _pos = savePos;
        }
    }

    private void ParseCmapSubtable(uint offset, ushort platformID, ushort encodingID)
    {
        _pos = (int)offset;
        var format = ReadUInt16();

        // 只处理常用的格式
        if (format == 4) // 分段映射
        {
            var length = ReadUInt16();
            var language = ReadUInt16();
            var segCountX2 = ReadUInt16();
            var segCount = segCountX2 / 2;
            var searchRange = ReadUInt16();
            var entrySelector = ReadUInt16();
            var rangeShift = ReadUInt16();

            var endCodes = new ushort[segCount];
            for (int i = 0; i < segCount; i++)
                endCodes[i] = ReadUInt16();

            var reservedPad = ReadUInt16();

            var startCodes = new ushort[segCount];
            for (int i = 0; i < segCount; i++)
                startCodes[i] = ReadUInt16();

            var idDeltas = new short[segCount];
            for (int i = 0; i < segCount; i++)
                idDeltas[i] = ReadInt16();

            var idRangeOffsets = new ushort[segCount];
            for (int i = 0; i < segCount; i++)
                idRangeOffsets[i] = ReadUInt16();

            // 构建映射
            for (int seg = 0; seg < segCount; seg++)
            {
                for (uint c = startCodes[seg]; c <= endCodes[seg]; c++)
                {
                    int glyphId;
                    if (idRangeOffsets[seg] == 0)
                    {
                        glyphId = (int)((c + (uint)idDeltas[seg]) & 0xFFFF);
                    }
                    else
                    {
                        var rangeOffset = idRangeOffsets[seg] / 2 + (int)(c - startCodes[seg]) - (segCount - seg);
                        glyphId = ReadUInt16At(_pos + rangeOffset * 2);
                        if (glyphId != 0)
                            glyphId = (glyphId + idDeltas[seg]) & 0xFFFF;
                    }

                    if (glyphId != 0)
                    {
                        var ch = (char)c;
                        GlyphToUnicode[glyphId] = ch;
                        UnicodeToGlyph[ch] = glyphId;
                    }
                }
            }
        }
        else if (format == 12) // 32位Unicode
        {
            var reserved = ReadUInt16();
            var length = ReadUInt32();
            var language = ReadUInt32();
            var numGroups = ReadUInt32();

            for (uint i = 0; i < numGroups; i++)
            {
                var startCharCode = ReadUInt32();
                var endCharCode = ReadUInt32();
                var startGlyphID = ReadUInt32();

                for (uint c = startCharCode; c <= endCharCode; c++)
                {
                    var glyphId = (int)(startGlyphID + (c - startCharCode));
                    var ch = (char)c;
                    GlyphToUnicode[glyphId] = ch;
                    UnicodeToGlyph[ch] = glyphId;
                }
            }
        }
    }

    // 获取字符的字形ID
    public int GetGlyphId(char c)
    {
        return UnicodeToGlyph.TryGetValue(c, out var glyphId) ? glyphId : 0;
    }

    // 获取字形宽度
    public int GetGlyphWidth(int glyphId)
    {
        if (GlyphWidths == null || GlyphWidths.Length == 0) return UnitsPerEm / 2;
        if (glyphId < GlyphWidths.Length) return GlyphWidths[glyphId];
        return GlyphWidths[^1]; // Last metric applies to all subsequent glyphs
    }

    // 读取辅助方法
    private byte[] ReadBytes(int count)
    {
        var result = new byte[count];
        Array.Copy(_data, _pos, result, 0, count);
        _pos += count;
        return result;
    }

    private byte ReadByte() => _data[_pos++];
    private ushort ReadUInt16() => (ushort)((ReadByte() << 8) | ReadByte());
    private short ReadInt16() => (short)ReadUInt16();
    private uint ReadUInt32() => ((uint)ReadUInt16() << 16) | ReadUInt16();
    private ulong ReadUInt64() => ((ulong)ReadUInt32() << 32) | ReadUInt32();

    private ushort ReadUInt16At(int position)
    {
        return (ushort)((_data[position] << 8) | _data[position + 1]);
    }

    private string ReadTag()
    {
        var bytes = ReadBytes(4);
        return Encoding.ASCII.GetString(bytes);
    }
}
