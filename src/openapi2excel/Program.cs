using ClosedXML.Excel;
using Microsoft.OpenApi.Readers;
using System.CommandLine;

namespace scl;

class Program
{
    static async Task<int> Main(string[] args)
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
        rootCommand.SetHandler(async (inFile, outFile) =>
            {
                var openApiReadResult = await ReadFile(inFile!);

                using var workbook = new XLWorkbook();
                foreach (var openApiPath in openApiReadResult.OpenApiDocument.Paths)
                {
                    foreach (var openApiOperation in openApiPath.Value.Operations)
                    {
                        var worksheet = workbook.Worksheets.Add(openApiOperation.Value.OperationId);
                        worksheet.Cell("A1").Value = openApiPath.Key;
                        worksheet.Cell("A2").Value = openApiOperation.Key.ToString();
                        worksheet.Cell("A3").Value = openApiOperation.Value.Summary;
                    }
                }

                workbook.SaveAs(outFile!.FullName);
            },
            inputFileOption, outputFileOption);

        return await rootCommand.InvokeAsync(args);
    }

    static async Task<ReadResult> ReadFile(FileInfo file)
    {
        await using var fileStream = file.OpenRead();
        var readResult = await new OpenApiStreamReader().ReadAsync(fileStream);
        return readResult;
    }
}