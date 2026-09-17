using System.Text;
using ClosedXML.Excel;
using openapi2excel.core;
using openapi2excel.core.Sanitization;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// Generated specifications violate the standard in ways that a human reader does not even
   /// notice: a renamed route parameter leaving two paths with the same signature, a typo between
   /// the path template and the parameter list, a parameter left behind as "in: path". The reader
   /// rejects such a document, the sanitizer repairs it and says what it changed.
   /// </summary>
   public class DocumentSanitizationTest
   {
      /// <summary>
      /// One instance of every violation the sanitizer repairs, modelled on a real Swashbuckle
      /// document: duplicated signature with disjoint methods, a case typo in a path parameter
      /// name, and a parameter that is not in the path template at all.
      /// </summary>
      private const string RepairableDocument = """
         openapi: 3.0.1
         info:
           title: Broken specification
           version: v1
         paths:
           /pets/{pet-id}:
             get:
               operationId: GetPet
               parameters:
                 - { name: pet-id, in: path, required: true, schema: { type: string } }
               responses:
                 '200': { description: OK }
           /pets/{petId}:
             delete:
               operationId: DeletePet
               parameters:
                 - { name: petId, in: path, required: true, schema: { type: string } }
               responses:
                 '200': { description: OK }
           /owners/{id}:
             delete:
               operationId: DeleteOwner
               parameters:
                 - { name: Id, in: path, required: true, schema: { type: string } }
               responses:
                 '200': { description: OK }
           /owners/{ownerId}/validate:
             get:
               operationId: ValidateOwner
               parameters:
                 - { name: ownerId, in: path, required: true, schema: { type: string } }
                 - { name: validationRole, in: path, required: true, schema: { type: string } }
               responses:
                 '200': { description: OK }
         """;

      /// <summary>Both paths of the duplicated signature define GET, so nothing can be merged.</summary>
      private const string UnrepairableDocument = """
         openapi: 3.0.1
         info:
           title: Broken specification
           version: v1
         paths:
           /pets/{pet-id}:
             get:
               operationId: GetPetById
               parameters:
                 - { name: pet-id, in: path, required: true, schema: { type: string } }
               responses:
                 '200': { description: OK }
           /pets/{petId}:
             get:
               operationId: GetPet
               parameters:
                 - { name: petId, in: path, required: true, schema: { type: string } }
               responses:
                 '200': { description: OK }
         """;

      [Fact]
      public async Task Repairable_document_is_generated_and_every_correction_is_reported()
      {
         var (workbook, report) = await GenerateAsync(RepairableDocument);
         using (workbook)
         {
            Assert.NotNull(report);
            Assert.Empty(report!.SkippedProblems);
            Assert.Equal(
            [
               SanitizationRule.DuplicatePathSignature,
               SanitizationRule.PathParameterRenamedToTemplate,
               SanitizationRule.PathParameterMovedToQuery
            ], report.Corrections.Select(correction => correction.Rule));

            Assert.Equal("/pets/{petId}", report.Corrections[0].Location);
            Assert.Contains("DELETE moved to '/pets/{pet-id}'", report.Corrections[0].Description);
            Assert.Contains("'petId' renamed to 'pet-id'", report.Corrections[0].Description);

            Assert.Equal("DELETE /owners/{id}", report.Corrections[1].Location);
            Assert.Contains("'Id' renamed to 'id'", report.Corrections[1].Description);

            Assert.Equal("GET /owners/{ownerId}/validate", report.Corrections[2].Location);
            Assert.Contains("'validationRole' is not in the path template", report.Corrections[2].Description);
         }
      }

      [Fact]
      public async Task Merged_path_keeps_every_operation_of_both_paths()
      {
         var (workbook, _) = await GenerateAsync(RepairableDocument);
         using (workbook)
         {
            Assert.Equal(["Info", "GetPet", "DeletePet", "DeleteOwner", "ValidateOwner"],
               workbook.Worksheets.Select(worksheet => worksheet.Name));

            // Both operations of the merged path are documented under the surviving path.
            Assert.Equal("/pets/{pet-id}", ValueRightOf(workbook.Worksheet("GetPet"), "Path"));
            Assert.Equal("/pets/{pet-id}", ValueRightOf(workbook.Worksheet("DeletePet"), "Path"));
         }
      }

      [Fact]
      public async Task Corrected_parameters_land_in_the_generated_document()
      {
         var (workbook, _) = await GenerateAsync(RepairableDocument);
         using (workbook)
         {
            // Renamed to match the template, still a path parameter.
            Assert.Equal(["id", "PATH", "Simple", "string", "Yes", "No", "No"],
               ParameterRow(workbook.Worksheet("DeleteOwner"), "id"));

            // Not in the template, so documented as an optional query parameter.
            Assert.Equal(["ownerId", "PATH", "Simple", "string", "Yes", "No", "No"],
               ParameterRow(workbook.Worksheet("ValidateOwner"), "ownerId"));
            Assert.Equal(["validationRole", "QUERY", "Form", "string", "No", "No", "No"],
               ParameterRow(workbook.Worksheet("ValidateOwner"), "validationRole"));
         }
      }

      [Fact]
      public async Task Disabled_sanitizer_rejects_the_document_and_reports_nothing()
      {
         SanitizationReport? report = null;
         var options = new OpenApiDocumentationOptions
         {
            SanitizeDocument = false,
            OnDocumentSanitized = sanitized => report = sanitized
         };

         var exception = await Assert.ThrowsAsync<InvalidOperationException>(
            () => GenerateAsync(RepairableDocument, options));

         Assert.Contains("Some errors occurred while processing input file.", exception.Message);
         Assert.Contains("The path signature '/pets/{}' MUST be unique.", exception.Message);
         Assert.Null(report);
      }

      [Fact]
      public async Task Duplicated_signature_with_the_same_method_is_reported_and_the_document_is_rejected()
      {
         SanitizationReport? report = null;
         var options = new OpenApiDocumentationOptions { OnDocumentSanitized = sanitized => report = sanitized };

         var exception = await Assert.ThrowsAsync<InvalidOperationException>(
            () => GenerateAsync(UnrepairableDocument, options));

         // Merging would drop one of the two GET operations, so the document stays invalid.
         Assert.Contains("The path signature '/pets/{}' MUST be unique.", exception.Message);
         Assert.NotNull(report);
         Assert.Empty(report!.Corrections);
         var problem = Assert.Single(report.SkippedProblems);
         Assert.Equal(SanitizationRule.DuplicatePathSignature, problem.Rule);
         Assert.Equal("/pets/{petId}", problem.Location);
         Assert.Contains("both define GET", problem.Description);
      }

      [Fact]
      public async Task Valid_document_is_left_alone()
      {
         SanitizationReport? report = null;
         var outputFile = Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");
         try
         {
            await using var input = File.OpenRead(Path.Combine("Sample", "Sample1.yaml"));
            await OpenApiDocumentationGenerator.GenerateDocumentation(input, outputFile,
               new OpenApiDocumentationOptions { OnDocumentSanitized = sanitized => report = sanitized });

            Assert.Null(report);
         }
         finally
         {
            File.Delete(outputFile);
         }
      }

      private static async Task<(XLWorkbook Workbook, SanitizationReport? Report)> GenerateAsync(string document,
         OpenApiDocumentationOptions? options = null)
      {
         SanitizationReport? report = null;
         if (options is null)
         {
            options = new OpenApiDocumentationOptions { OnDocumentSanitized = sanitized => report = sanitized };
         }
         else
         {
            var caller = options.OnDocumentSanitized;
            options.OnDocumentSanitized = sanitized =>
            {
               report = sanitized;
               caller?.Invoke(sanitized);
            };
         }

         var outputFile = Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");
         try
         {
            using var input = new MemoryStream(Encoding.UTF8.GetBytes(document));
            await OpenApiDocumentationGenerator.GenerateDocumentation(input, outputFile, options);

            await using var generated = File.OpenRead(outputFile);
            return (new XLWorkbook(generated), report);
         }
         finally
         {
            File.Delete(outputFile);
         }
      }

      /// <summary>The value of a labelled row of the operation information block.</summary>
      private static string ValueRightOf(IXLWorksheet worksheet, string label)
      {
         var cell = worksheet.CellsUsed(cell => cell.GetString() == label).Single();
         return worksheet.Row(cell.Address.RowNumber).CellsUsed()
            .Last(used => used.Address.ColumnNumber > cell.Address.ColumnNumber).GetString();
      }

      /// <summary>Every filled cell of the parameter row with the given name.</summary>
      private static List<string> ParameterRow(IXLWorksheet worksheet, string name)
      {
         var cell = worksheet.CellsUsed(cell => cell.GetString() == name).Single();
         return worksheet.Row(cell.Address.RowNumber).CellsUsed()
            .Select(used => used.GetString())
            .Where(text => !string.IsNullOrEmpty(text))
            .ToList();
      }
   }
}
