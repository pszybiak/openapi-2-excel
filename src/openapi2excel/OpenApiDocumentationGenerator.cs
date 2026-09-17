using ClosedXML.Excel;
using Microsoft.OpenApi.Extensions;
using Microsoft.OpenApi.Models;
using Microsoft.OpenApi.Readers;
using Microsoft.OpenApi.Validations;
using openapi2excel.core.Builders;
using openapi2excel.core.Common;
using openapi2excel.core.Sanitization;
using System.Text;

namespace openapi2excel.core;

public static class OpenApiDocumentationGenerator
{
   public static async Task GenerateDocumentation(string openApiFile, string outputFile)
   {
      if (!File.Exists(openApiFile))
         throw new FileNotFoundException($"Invalid input file path: {openApiFile}.");

      if (string.IsNullOrEmpty(outputFile))
         throw new ArgumentNullException(outputFile, "Invalid output file path.");

      await using var fileStream = File.OpenRead(openApiFile);
      await GenerateDocumentationImpl(fileStream, outputFile, new OpenApiDocumentationOptions());
   }

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
      AssertDocumentIsUsable(readResult, options);
      UnresolvedReferences.ResolveResponseHeaderSchemas(readResult.OpenApiDocument);

      using var workbook = new XLWorkbook();
      var infoWorksheetsBuilder = new InfoWorksheetBuilder(workbook, options);
      infoWorksheetsBuilder.Build(readResult.OpenApiDocument);

      // The Object type column of every operation names a schema the last worksheet documents, and
      // that worksheet does not exist yet, so the links are collected here and set at the end.
      var objectLinks = new ObjectLinkRegistry();
      var worksheetBuilder = new OperationWorksheetBuilder(workbook, options, objectLinks);
      readResult.OpenApiDocument.Paths.ForEach(path
         => path.Value.Operations.ForEach(operation
               =>
               {
                  var worksheet = worksheetBuilder.Build(path.Key, path.Value, operation.Key, operation.Value);
                  infoWorksheetsBuilder.AddLink(path.Key, path.Value, operation.Key, operation.Value, worksheet);
               }
         ));

      // Every worksheet exists by now, so the index can list them grouped instead of in the order
      // the document happens to declare its paths.
      infoWorksheetsBuilder.AddOperationsIndex();

      // The objects come last, after every worksheet naming one, and linking them is the last thing
      // the document needs.
      var objectsWorksheet = new ObjectsWorksheetBuilder(workbook, options, objectLinks)
         .Build(ObjectsCatalog.Collect(readResult.OpenApiDocument));
      if (objectsWorksheet is not null)
      {
         infoWorksheetsBuilder.AddObjectsLink(objectsWorksheet);
      }

      objectLinks.Resolve();

      workbook.SaveAs(new FileInfo(outputFile).FullName);
   }

   private static void AssertDocumentIsUsable(ReadResult readResult, OpenApiDocumentationOptions options)
   {
      var errors = readResult.OpenApiDiagnostic.Errors.ToList();
      if (!errors.Any())
         return;

      if (options.SanitizeDocument)
      {
         var report = OpenApiDocumentSanitizer.Sanitize(readResult.OpenApiDocument);
         if (report.HasChanges)
         {
            options.OnDocumentSanitized?.Invoke(report);
         }

         // The validator reports the same violations as the reader, so this is what is left of them.
         errors = Validate(readResult.OpenApiDocument);
         if (!errors.Any())
            return;
      }

      var errorMessageBuilder = new StringBuilder();
      errorMessageBuilder.AppendLine("Some errors occurred while processing input file.");
      errors.ForEach(e => errorMessageBuilder.AppendLine($"{e.Message} ({e.Pointer})"));
      throw new InvalidOperationException(errorMessageBuilder.ToString());
   }

   private static List<OpenApiError> Validate(OpenApiDocument document)
      => document.Validate(ValidationRuleSet.GetDefaultRuleSet()).ToList();
}