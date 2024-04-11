using openapi2excel.core;

namespace OpenApi2Excel.Tests
{
   public class OpenApiDocumentationGeneratorTest
   {
      [Fact]
      public async Task GenerateDocumentation_create_excel_file_for_correct_openapi_document()
      {
         const string inputFIle = "Sample/Sample1.yaml";
         const string outputFile = "output.xlsx";

         await using var file = File.OpenRead(inputFIle);

         await OpenApiDocumentationGenerator.GenerateDocumentation(file, outputFile, new OpenApiDocumentationOptions());

         Assert.True(File.Exists(outputFile));
      }
   }
}