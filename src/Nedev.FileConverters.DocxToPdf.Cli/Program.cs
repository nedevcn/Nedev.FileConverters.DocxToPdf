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
    // use the shared Converter entry point so that other converters can be used
    using var inStream = File.OpenRead(inputPath);
    using var result = Converter.Convert(inStream, "docx", "pdf");
    using var outStream = File.Create(outputPath);
    result.CopyTo(outStream);

    Console.WriteLine($"Converted '{inputPath}' -> '{outputPath}' successfully.");
}
catch (Exception ex)
{
    Console.Error.WriteLine("Conversion failed: " + ex.Message);
    Environment.Exit(1);
}
