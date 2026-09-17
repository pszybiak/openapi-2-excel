using ClosedXML.Excel;
using openapi2excel.core;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// Content gate for the generated workbook. A "file exists" assertion cannot see a dependency
   /// upgrade (ClosedXML, Microsoft.OpenApi) silently changing what lands in the cells - these can.
   /// Expected values come from Sample/Sample1.yaml (Swagger Petstore 1.0.11).
   /// </summary>
   public class GeneratedWorkbookContentTest : IClassFixture<GeneratedWorkbookContentTest.GeneratedWorkbook>
   {
      private readonly XLWorkbook _workbook;

      public GeneratedWorkbookContentTest(GeneratedWorkbook workbook) => _workbook = workbook.Workbook;

      [Fact]
      public void Workbook_contains_info_worksheet_one_worksheet_per_operation_and_the_objects()
      {
         string[] expected =
         [
            "Info", "updatePet", "addPet", "findPetsByStatus", "findPetsByTags", "getPetById",
            "updatePetWithForm", "deletePet", "uploadFile", "getInventory", "placeOrder",
            "getOrderById", "deleteOrder", "createUser", "createUsersWithListInput", "loginUser",
            "logoutUser", "getUserByName", "updateUser", "deleteUser", "Objects"
         ];

         Assert.Equal(expected, _workbook.Worksheets.Select(worksheet => worksheet.Name));
      }

      [Fact]
      public void Info_worksheet_describes_the_api_and_links_to_every_operation()
      {
         var info = _workbook.Worksheet("Info");

         Assert.Equal("Version", info.Cell("C1").GetString());
         Assert.Equal("1.0.11", info.Cell("D1").GetString());
         Assert.Equal("Title", info.Cell("C2").GetString());
         Assert.Equal("Swagger Petstore - OpenAPI 3.0", info.Cell("D2").GetString());

         // The index of the operations starts below the description of the api.
         Assert.Equal("OPERATIONS", info.Cell("A12").GetString());
         Assert.Equal(["Tag", "Path", "Operation type: Id", "Operation summary"],
            Enumerable.Range(1, 4).Select(column => info.Cell(13, column).GetString()));

         // Groups of the specification, in the order it declares its tags.
         Assert.Equal(["pet", "store", "user"],
            new[] { 14, 28, 36 }.Select(row => info.Cell(row, 1).GetString()));
         Assert.Equal("Everything about your Pets", info.Cell("D14").GetString());

         // Path of the tag, then the operations of the path, each linking to its worksheet. One
         // cell states the type of the operation and the id the specification names it with.
         Assert.Equal("/pet", info.Cell("B15").GetString());
         Assert.Equal("PUT: updatePet", info.Cell("C16").GetString());
         Assert.Equal("Update an existing pet", info.Cell("D16").GetString());
         Assert.Equal("'updatePet'!A1", info.Cell("C16").GetHyperlink().InternalAddress);
         Assert.Equal("'updatePet'!A1", info.Cell("D16").GetHyperlink().InternalAddress);
         Assert.Equal([0, 1, 2], new[] { 14, 15, 16 }.Select(row => info.Row(row).OutlineLevel));

         // 19 operations in the sample, each carrying one tag, so each is listed once.
         Assert.Equal(19, info.Rows(14, 48).Count(row => row.Cell(3).HasHyperlink));

         // The link to the worksheet documenting the objects closes the index.
         Assert.Equal("OBJECTS", info.Cell("A50").GetString());
         Assert.Equal("'Objects'!A1", info.Cell("A50").GetHyperlink().InternalAddress);
         Assert.Equal(50, info.LastRowUsed()!.RowNumber());
      }

      [Fact]
      public void Operation_worksheet_describes_the_operation_and_links_back_to_the_info_worksheet()
      {
         var operation = _workbook.Worksheet("updatePet");

         Assert.Equal("<<<<<", operation.Cell("A1").GetString());
         Assert.Equal("'Info'!A1", operation.Cell("A1").GetHyperlink().InternalAddress);

         Assert.Equal("OPERATION INFORMATION", operation.Cell("A3").GetString());
         Assert.True(operation.Cell("A3").Style.Font.Bold);
         Assert.Equal("PUT", operation.Cell("E4").GetString());
         Assert.Equal("updatePet", operation.Cell("E5").GetString());
         Assert.Equal("/pet", operation.Cell("E6").GetString());
         Assert.Equal("Update an existing pet by Id", operation.Cell("E7").GetString());
         Assert.Equal("Update an existing pet", operation.Cell("E8").GetString());
         Assert.Equal("No", operation.Cell("E9").GetString());
      }

      [Fact]
      public void Schema_table_header_is_bold_and_filled_and_carries_every_column()
      {
         var operation = _workbook.Worksheet("updatePet");
         string[] expected =
         [
            "Name", "", "", "", "", "Type", "Object type", "Format", "Length", "Required",
            "Nullable", "Range", "Pattern", "Enum", "Deprecated", "Default", "Example", "Description"
         ];

         Assert.Equal(expected, Enumerable.Range(1, 18).Select(column => operation.Row(13).Cell(column).GetString()));
         Assert.True(operation.Cell("A13").Style.Font.Bold);
         Assert.True(operation.Cell("R13").Style.Font.Bold);
         Assert.Equal(XLColor.LightGray, operation.Cell("A13").Style.Fill.BackgroundColor);
      }

      [Fact]
      public void Schema_table_renders_types_formats_examples_enums_and_nesting()
      {
         var operation = _workbook.Worksheet("updatePet");

         // Flat property: name in column A, type/format/required/nullable/example across the row.
         Assert.Equal("id", operation.Cell("A14").GetString());
         Assert.Equal("integer", operation.Cell("F14").GetString());
         Assert.Equal("int64", operation.Cell("H14").GetString());
         Assert.Equal("No", operation.Cell("J14").GetString());
         Assert.Equal("No", operation.Cell("K14").GetString());
         Assert.Equal("10", operation.Cell("Q14").GetString());

         // Required flag comes from the schema, not from the property itself.
         Assert.Equal("name", operation.Cell("A15").GetString());
         Assert.Equal("Yes", operation.Cell("J15").GetString());

         // Referenced object: type "object" plus the reference id.
         Assert.Equal("category", operation.Cell("A16").GetString());
         Assert.Equal("object", operation.Cell("F16").GetString());
         Assert.Equal("Category", operation.Cell("G16").GetString());

         // Nested properties are indented one column to the right.
         Assert.Equal("id", operation.Cell("B17").GetString());
         Assert.Equal("", operation.Cell("A17").GetString());

         // Array of primitives renders a synthetic <value> item row.
         Assert.Equal("photoUrls", operation.Cell("A19").GetString());
         Assert.Equal("array", operation.Cell("F19").GetString());
         Assert.Equal("<value>", operation.Cell("B20").GetString());
         Assert.Equal("string", operation.Cell("F20").GetString());

         // Enum values are joined, description is copied.
         Assert.Equal("status", operation.Cell("A24").GetString());
         Assert.Equal("available, pending, sold", operation.Cell("N24").GetString());
         Assert.Equal("pet status in the store", operation.Cell("R24").GetString());
      }

      [Fact]
      public void Operation_worksheet_documents_every_body_format_and_response()
      {
         var operation = _workbook.Worksheet("updatePet");
         var texts = operation.CellsUsed().Select(cell => cell.GetString()).ToList();

         Assert.Equal("REQUEST", operation.Cell("A11").GetString());

         // Every body of the operation is a Pet, and the title of a body format says which object
         // it is when the specification gives that object a name.
         Assert.Equal(2, texts.Count(text => text == "Body format: application/json (Pet)"));
         Assert.Equal(2, texts.Count(text => text == "Body format: application/xml (Pet)"));
         Assert.Single(texts, text => text == "Body format: application/x-www-form-urlencoded (Pet)");
         Assert.Contains("RESPONSE", texts);

         // Guards against silent content loss: the sample operation fills 87 rows.
         Assert.Equal(87, operation.LastRowUsed()!.RowNumber());
      }

      public sealed class GeneratedWorkbook : IDisposable
      {
         private readonly string _outputFile =
            Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");

         public GeneratedWorkbook()
         {
            using var input = File.OpenRead(Path.Combine("Sample", "Sample1.yaml"));
            OpenApiDocumentationGenerator
               .GenerateDocumentation(input, _outputFile, new OpenApiDocumentationOptions())
               .GetAwaiter().GetResult();

            Workbook = new XLWorkbook(_outputFile);
         }

         public XLWorkbook Workbook { get; }

         public void Dispose()
         {
            Workbook.Dispose();
            File.Delete(_outputFile);
         }
      }
   }
}
