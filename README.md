# DocxToPdf

A high-performance .NET 10 library and CLI tool for converting DOCX files to PDF using a custom-built PDF engine.

## 🌟 Features

- **Rich Document Support**: Converts complex DOCX elements including:
    - Paragraphs with advanced styling (indentation, spacing, alignment).
    - Multi-level lists and numbering.
    - Tables with cell merging, borders, and custom background colors.
    - Images (Inline and Floating) with support for text wrapping (Square, Tight, Through, Top/Bottom).
    - Hyperlinks and Bookmarks.
- **Advanced Layout**:
    - Multi-column page layouts.
    - Section-specific page settings (size, margins, orientation).
    - Headers and Footers (Different first page, even/odd support).
- **Professional Enhancements**:
    - Dynamic Watermarks (Text-based).
    - Document Field Resolution (DATE, AUTHOR, TITLE, etc.).
    - Footnotes and Endnotes support.
    - Comments Summary Page generation.
- **Custom PDF Engine**: Built on top of a specialized PDF generation layer for precise control over rendering.

## 🛠 Technical Stack

- **Framework**: .NET 10
- **Core Dependencies**:
    - `DocumentFormat.OpenXml`: For robust DOCX parsing.
    - `SkiaSharp`: For high-quality text and image rendering.
- **Architecture**: Modular converter design with dedicated processors for paragraphs, tables, images, and lists.

## 🚀 Getting Started

### Prerequisites

- .NET 10 SDK

### Installation

Clone the repository and build the project:

```bash
git clone <repository-url>
cd DocxToPdf/src
dotnet build
```

### Usage (CLI)

Run the converter from the command line:

```bash
dotnet run -- <input.docx> <output.pdf>
```

If no arguments are provided, the tool defaults to `test.docx` as input and `test.pdf` as output.

## 📂 Project Structure

- `src/DocxToPdfConverter.cs`: The main orchestration logic for the conversion process.
- `src/PdfEngine/`: Core PDF generation and rendering engine.
- `src/Converters/`: Specialized logic for handling different DOCX elements (Tables, Images, etc.).
- `src/Helpers/`: Utility classes for font handling, styling, and OpenXml extensions.
- `src/Models/`: Data models used during the conversion process.

## 📄 License

[Insert License Information Here]
