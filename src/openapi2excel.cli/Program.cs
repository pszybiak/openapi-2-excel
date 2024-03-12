using System.CommandLine;

namespace OpenApi2Excel.cli;

internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        var inputFileOption = new Option<FileInfo?>(
            name: "--file",
            description: "The input file to read.");

        var outputFileOption = new Option<FileInfo?>(
            name: "--out",
            description: "The file to write.");

        var rootCommand = new RootCommand("OpenApi-2-Excel");
        rootCommand.AddOption(inputFileOption);
        rootCommand.AddOption(outputFileOption);
        rootCommand.SetHandler(async (inFile, outFile)
                => await OpenApiDocumentationGenerator.GenerateDocumentation(inFile!.FullName, outFile!.FullName, new OpenApiDocumentationOptions()).ConfigureAwait(false),
            inputFileOption, outputFileOption);

        return await rootCommand.InvokeAsync(args);
    }
}