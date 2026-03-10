# DocxToPdf

A high-performance .NET 10 library and CLI tool for converting DOCX files to PDF using a custom-built PDF engine.

## 🌟 Features

- **Rich Document Support**: Converts complex DOCX elements including:
    - Paragraphs with advanced styling (indentation, spacing, alignment).
    - Multi-level lists and numbering.
    - Tables with cell merging, borders, and custom background colors.
    - Images (Inline and Floating) with support for text wrapping (Square, Tight, Through, Top/Bottom).
    - Hyperlinks and Bookmarks (both <w:hyperlink> elements and HYPERLINK fields preserve display text and produce clickable links).
    - Basic handling of complex fields (using FieldChar/FieldCode) – common cases like PAGE, DATE, REF, MERGEFIELD, and HYPERLINK are resolved.
- **Advanced Layout & Styling**:
    - Page Borders (pgBorders) with support for custom colors, sizes, and offsets (Page/Text).
    - Basic VML Shape support (WordArt, Rect, Oval) with fill and stroke styling.
    - Multi-column page layouts and section-specific settings.
    - Headers and Footers (Different first page, even/odd support).
- **Professional Enhancements**:
    - Dynamic Watermarks (Text-based).
    - Document Field Resolution (DATE, AUTHOR, TITLE, etc.).
    - Footnotes and Endnotes support.
    - Comments Summary Page generation.
- **Custom PDF Engine**: Built on top of a specialized PDF generation layer for precise control over rendering.

## 📈 Feature Completeness & Roadmap

**Current State:** The library robustly handles standard document layouts including paragraphs, multi-level lists, nested tables, headers/footers, and images (inline/floating). 

**Next Steps / Roadmap:**
1. **Test Coverage Expansion:** Introduce comprehensive unit tests for `TableConverter`, `ListConverter`, `ImageConverter`, and the core `PdfEngine`.
2. **Advanced Field Semantics:** Support fully nested OpenXML field codes and dynamic condition evaluation.
3. **DrawingML / Math:** Add proper vector rendering for complex shapes (`<w:pict>`, `<wp:inline>`) and OMML math blocks.
4. **Performance Tuning:** Audit memory management and stream processing for massive document (1000+ pages) conversions.

## 🛠 Technical Stack
> **Field support:** The converter recognises both simple fields (`<w:fldSimple>`) and
> complex fields built from `FieldChar/FieldCode` runs. Common field codes such as
> `PAGE`, `DATE`, `REF`, `MERGEFIELD` and `HYPERLINK` are resolved; the display
> text from complex fields is preserved when available. Full OpenXML field
> semantics (nested fields, switching before separator, etc.) are a future
> enhancement.

- **Framework**: .NET 10
- **Core Dependencies**:
    - `DocumentFormat.OpenXml`: For robust DOCX parsing.
    - `SkiaSharp`: For high-quality text and image rendering.
    - `Nedev.FileConverters.Core` (>=0.1.0): shared converter interfaces and discovery
- **Architecture**: Modular converter design with dedicated processors for paragraphs, tables, images, and lists.

## 🚀 Getting Started

### Prerequisites

- .NET 10 SDK

### Installation

Clone the repository and build the project:

```bash
git clone <repository-url>
cd Nedev.FileConverters.DocxToPdf/src
dotnet build
```

### Usage (CLI)

There is now a dedicated console application under `src/Nedev.FileConverters.DocxToPdf.Cli`.

You can invoke it using the CLI project:

```bash
cd Nedev.FileConverters.DocxToPdf/src/Nedev.FileConverters.DocxToPdf.Cli
dotnet run -- <input.docx> <output.pdf>
```

When run without arguments, it defaults to `test.docx` → `test.pdf`.

The CLI also demonstrates how to call the shared `Nedev.FileConverters.Converter` entry point.  In a host application you can simply load this assembly (or reference the project) and use:

```csharp
using Nedev.FileConverters;

using var outStream = Converter.Convert(inStream, "docx", "pdf");
```

Because `DocxToPdfConverter` is attributed, the core library will discover and instantiate it automatically when a conversion is requested.

For applications using dependency injection you can register the converter in two ways:

```csharp
// using core method directly:
services.AddFileConverter("docx", "pdf", new DocxToPdfConverter());

// or via the convenience helper provided by this library:
services.AddDocxToPdf();
```


## 📂 Project Structure

- `src/Nedev.FileConverters.DocxToPdfConverter.cs`: The main orchestration logic for the conversion process. Implements `IFileConverter` and is decorated with `[FileConverter("docx","pdf")]` so it is automatically discovered by the core library.
- `src/PdfEngine/`: Core PDF generation and rendering engine.
- `src/Converters/`: Specialized logic for handling different DOCX elements (Tables, Images, etc.).
- `src/Helpers/`: Utility classes for font handling, styling, and OpenXml extensions.
- `src/Models/`: Data models used during the conversion process (now extends/works alongside core models if applicable).
- `src/Nedev.FileConverters.DocxToPdf.Cli`: Simple console application demonstrating both direct and `Nedev.FileConverters.Converter` usage.

## 📄 License

This project is released under the [MIT License](../LICENSE).

The same license is included in the NuGet package via `PackageLicenseExpression`.
