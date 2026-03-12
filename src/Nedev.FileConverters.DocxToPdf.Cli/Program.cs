using Nedev.FileConverters.DocxToPdf;
using Nedev.FileConverters;

// simple CLI wrapper for the DocxToPdfConverter / core infrastructure
var argsList = args.Length == 0
    ? new[] { "test.docx", "test.pdf" }
    : args;

if (argsList.Length < 2)
{
    Console.WriteLine("Usage: dotnet run -- <input.docx> <output.pdf>");
    Console.WriteLine("The converter is also discoverable via the Nedev.FileConverters.Core infrastructure.");
    return;
}

var inputPath = argsList[0];
var outputPath = argsList[1];

// Validate input file exists
if (!File.Exists(inputPath))
{
    Console.Error.WriteLine($"Error: Input file not found: '{inputPath}'");
    Environment.Exit(1);
    return;
}

// Validate input file extension
if (!inputPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) &&
    !inputPath.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine($"Warning: Input file does not have a .docx or .doc extension: '{inputPath}'");
}

// Ensure output directory exists
var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
{
    try
    {
        Directory.CreateDirectory(outputDir);
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Error: Failed to create output directory '{outputDir}': {ex.Message}");
        Environment.Exit(1);
        return;
    }
}

try
{
    // call the DocxToPdf converter directly; avoid the generic infrastructure
    // which is useful in library scenarios but not needed for the CLI.
    // static helpers make a one‑liner possible.
    DocxToPdfConverter.ConvertFile(inputPath, outputPath);

    Console.WriteLine($"Converted '{inputPath}' -> '{outputPath}' successfully.");
}
catch (Exception ex)
{
    // log the entire exception object so inner details and stack trace are shown
    Console.Error.WriteLine("Conversion failed: " + ex);
    Environment.Exit(1);
}
