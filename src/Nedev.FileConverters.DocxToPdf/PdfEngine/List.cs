namespace Nedev.FileConverters.DocxToPdf.PdfEngine;

/// <summary>
/// ???
/// </summary>
public class ListItem : Paragraph
{
    public Chunk ListSymbol { get; set; } = new("�", Font.Helvetica(12));

    public ListItem(string? text = null, Font? font = null) : base(text, font) { }
}

/// <summary>
/// ??
/// </summary>
public class List : IElement
{
    public const int ORDERED = 0;
    public const int UNORDERED = 1;
    public const int ALPHABETICAL = 2;
    public const int ROMAN = 3;

    public int Type => 30;
    public bool IsContent() => true;
    public bool IsNestable() => false;

    public int ListType { get; set; } = UNORDERED;
    public float IndentationLeft { get; set; } = 20f;
    public float SymbolIndent { get; set; } = 15f;
    public bool Autoindent { get; set; } = true;

    public Chunk ListSymbol { get; set; } = new("�", Font.Helvetica(12));

    private readonly List<ListItem> _items = [];

    public List(int type = UNORDERED)
    {
        ListType = type;
        UpdateSymbol();
    }

    private void UpdateSymbol()
    {
        ListSymbol = ListType switch
        {
            ORDERED => new Chunk("1.", Font.Helvetica(12)),
            UNORDERED => new Chunk("�", Font.Helvetica(12)),
            ALPHABETICAL => new Chunk("a.", Font.Helvetica(12)),
            ROMAN => new Chunk("i.", Font.Helvetica(12)),
            _ => new Chunk("�", Font.Helvetica(12))
        };
    }

    public void Add(ListItem item)
    {
        _items.Add(item);
    }

    public void Add(string text)
    {
        _items.Add(new ListItem(text));
    }

    public IEnumerable<ListItem> Items => _items;

    public int Count => _items.Count;
}
