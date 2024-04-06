using ClosedXML.Excel;
using Microsoft.OpenApi.Readers;
using OpenApi2Excel.Builders;
using OpenApi2Excel.Common;
using System.Text;

namespace OpenApi2Excel;

public static class OpenApiDocumentationGenerator
{
   public static async Task GenerateDocumentation(string openApiFile, string outputFile,
      OpenApiDocumentationOptions options)
   {
      if (!File.Exists(openApiFile))
         throw new FileNotFoundException($"Invalid input file path: {openApiFile}.");

      if (string.IsNullOrEmpty(outputFile))
         throw new ArgumentNullException(outputFile, "Invalid output file path.");

      await using var fileStream = File.OpenRead(openApiFile);
      await GenerateDocumentationImpl(fileStream, outputFile, options);
   }

   public static async Task GenerateDocumentation(Stream openApiFileStream, string outputFile,
      OpenApiDocumentationOptions options)
   {
      if (string.IsNullOrEmpty(outputFile))
         throw new ArgumentNullException(outputFile, "Invalid output file path.");

      await GenerateDocumentationImpl(openApiFileStream, outputFile, options);
   }

   private static async Task GenerateDocumentationImpl(Stream openApiFileStream, string outputFile,
      OpenApiDocumentationOptions options)
   {
      var readResult = await new OpenApiStreamReader().ReadAsync(openApiFileStream);
      if (readResult.OpenApiDiagnostic.Errors.Any())
      {
         var errorMessageBuilder = new StringBuilder();
         errorMessageBuilder.AppendLine("Some errors occurred while processing input file.");
         readResult.OpenApiDiagnostic.Errors.ToList().ForEach(e => errorMessageBuilder.AppendLine(e.Message));
         throw new InvalidOperationException(errorMessageBuilder.ToString());
      }

      using var workbook = new XLWorkbook();
      var infoWorksheetsBuilder = new InfoWorksheetBuilder(workbook, options);
      var worksheetBuilder = new OperationWorksheetBuilder(workbook, options);

      infoWorksheetsBuilder.Build(readResult.OpenApiDocument);
      readResult.OpenApiDocument.Paths.ForEach(path
         => path.Value.Operations.ForEach(operation
               =>
            {
               var worksheet = worksheetBuilder.Build(path.Key, path.Value, operation.Key, operation.Value);
               infoWorksheetsBuilder.AddLink(operation.Key, path.Key, worksheet);
            }
         ));

      workbook.SaveAs(new FileInfo(outputFile).FullName);
   }
}