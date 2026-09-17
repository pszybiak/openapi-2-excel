using System.Text;
using ClosedXML.Excel;
using openapi2excel.core;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// Worksheet names come from operationId, which real specs repeat and make arbitrarily long,
   /// while Excel allows 31 characters and no duplicates.
   /// </summary>
   public class WorksheetNamingTest
   {
      private const string DuplicatedOperationIds = """
         openapi: 3.0.1
         info:
           title: Duplicated operation ids
           version: v1
         paths:
           /a/short:
             get:
               operationId: GetUser
               responses:
                 '200':
                   description: OK
           /b/short:
             get:
               operationId: GetUser
               responses:
                 '200':
                   description: OK
           /c/short:
             get:
               operationId: GetUser
               responses:
                 '200':
                   description: OK
           /a/long:
             get:
               operationId: GetTheVeryLongAndDetailedVerificationOfAnAsset
               responses:
                 '200':
                   description: OK
           /b/long:
             get:
               operationId: GetTheVeryLongAndDetailedVerificationOfAnAsset
               responses:
                 '200':
                   description: OK
           /no/operation/id:
             get:
               responses:
                 '200':
                   description: OK
         """;

      [Fact]
      public async Task Repeated_operation_id_shorter_than_the_name_limit_does_not_throw()
      {
         // Regression: the uniqueness suffix used to be built with name[..28], which threw
         // ArgumentOutOfRangeException for every duplicated operationId shorter than 28 characters.
         using var workbook = await GenerateAsync();

         Assert.Equal(["Info", "GetUser", "GetUser_2", "GetUser_3"],
            workbook.Worksheets.Select(worksheet => worksheet.Name).Take(4));
      }

      [Fact]
      public async Task Repeated_long_operation_id_is_trimmed_so_the_suffix_still_fits()
      {
         using var workbook = await GenerateAsync();
         var names = workbook.Worksheets.Select(worksheet => worksheet.Name).ToList();

         Assert.Contains("GetTheVeryLongAndDetailedVer", names);
         Assert.Contains("GetTheVeryLongAndDetailedV_2", names);
         Assert.All(names, name => Assert.True(name.Length <= 31, $"'{name}' is {name.Length} characters"));
      }

      [Fact]
      public async Task Operation_without_operation_id_is_named_after_method_and_path()
      {
         using var workbook = await GenerateAsync();

         Assert.Contains("GET_no-operation-id", workbook.Worksheets.Select(worksheet => worksheet.Name));
      }

      [Fact]
      public async Task Every_operation_gets_its_own_worksheet()
      {
         using var workbook = await GenerateAsync();

         // 6 operations plus the Info worksheet, all names unique.
         Assert.Equal(7, workbook.Worksheets.Count);
         Assert.Equal(7, workbook.Worksheets.Select(worksheet => worksheet.Name.ToUpperInvariant()).Distinct().Count());
      }

      private static async Task<XLWorkbook> GenerateAsync()
      {
         var outputFile = Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");
         try
         {
            using var input = new MemoryStream(Encoding.UTF8.GetBytes(DuplicatedOperationIds));
            await OpenApiDocumentationGenerator.GenerateDocumentation(input, outputFile,
               new OpenApiDocumentationOptions());

            // Load into memory so the file can be deleted right away.
            await using var generated = File.OpenRead(outputFile);
            return new XLWorkbook(generated);
         }
         finally
         {
            File.Delete(outputFile);
         }
      }
   }
}
