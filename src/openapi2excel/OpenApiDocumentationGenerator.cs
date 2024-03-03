using ClosedXML.Excel;
using Microsoft.OpenApi.Readers;

namespace openapi2excel;

public static class OpenApiDocumentationGenerator
{
    public static async Task GenerateDocumentation(string openApiFile, string outputFile, OpenApiDocumentationOptions options)
    {
        if (!File.Exists(openApiFile))
            throw new FileNotFoundException($"Invalid input file path: {openApiFile}.");

        if (string.IsNullOrEmpty(outputFile))
            throw new ArgumentNullException(outputFile, "Invalid output file path.");

        await using var fileStream = File.OpenRead(openApiFile);
        await GenerateDocumentationImpl(fileStream, outputFile, options);

    }

    public static async Task GenerateDocumentation(Stream openApiFileStream, string outputFile, OpenApiDocumentationOptions options)
    {
        if (string.IsNullOrEmpty(outputFile))
            throw new ArgumentNullException(outputFile, "Invalid output file path.");

        await GenerateDocumentationImpl(openApiFileStream, outputFile, options);
    }

    private static async Task GenerateDocumentationImpl(Stream openApiFileStream, string outputFile, OpenApiDocumentationOptions options)
    {
        var readResult = await new OpenApiStreamReader().ReadAsync(openApiFileStream);
        if (readResult.OpenApiDiagnostic.Errors.Any())
        {
            throw new InvalidOperationException("Some errors occurred while processing input file.");
        }

        using var workbook = new XLWorkbook();
        foreach (var openApiPath in readResult.OpenApiDocument.Paths)
        {
            foreach (var openApiOperation in openApiPath.Value.Operations)
            {
                var worksheet = workbook.Worksheets.Add(openApiOperation.Value.OperationId);
                worksheet.Cell("A1").Value = openApiPath.Key;
                worksheet.Cell("A2").Value = openApiOperation.Key.ToString();
                worksheet.Cell("A3").Value = openApiOperation.Value.Summary;
            }
        }

        workbook.SaveAs(new FileInfo(outputFile).FullName);
    }
}

public class OpenApiDocumentationOptions
{
}