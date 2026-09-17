using openapi2excel.core;

namespace OpenApi2Excel.Tests
{
   public class OpenApiDocumentationGeneratorTest
   {
      [Fact]
      public async Task GenerateDocumentation_create_excel_file_for_correct_openapi_document()
      {
         const string inputFile = "Sample/Sample1.yaml";

         // A fixed name in the output directory fails the gate as soon as anything still holds the
         // file of the last run, Excel itself while somebody looks at what the test produced.
         var outputFile = Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");
         try
         {
            await using var file = File.OpenRead(inputFile);

            await OpenApiDocumentationGenerator.GenerateDocumentation(file, outputFile,
               new OpenApiDocumentationOptions());

            Assert.True(File.Exists(outputFile));
         }
         finally
         {
            File.Delete(outputFile);
         }
      }
   }
}
