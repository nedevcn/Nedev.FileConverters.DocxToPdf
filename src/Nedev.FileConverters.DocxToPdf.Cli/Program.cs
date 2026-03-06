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
