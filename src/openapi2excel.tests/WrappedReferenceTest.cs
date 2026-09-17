using System.Text;
using ClosedXML.Excel;
using openapi2excel.core;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// A generator that has to add a keyword to a referenced schema, almost always "nullable: true",
   /// cannot put it next to the $ref, so it wraps the reference in a single element allOf. The name
   /// of the referenced type, and everything else the reference carries, has to survive that
   /// wrapper.
   /// </summary>
   public class WrappedReferenceTest : IClassFixture<WrappedReferenceTest.GeneratedWorkbook>
   {
      private readonly IXLWorksheet _worksheet;

      public WrappedReferenceTest(GeneratedWorkbook workbook) => _worksheet = workbook.Workbook.Worksheet("AddPet");

      [Fact]
      public void Reference_used_directly_keeps_working()
      {
         var property = PropertyRow("directReference");

         Assert.Equal("object", property["Type"]);
         Assert.Equal("CategoryDto", property["Object type"]);
         Assert.Equal("No", property["Nullable"]);
      }

      [Fact]
      public void Wrapped_reference_is_named_and_stays_nullable()
      {
         var property = PropertyRow("wrappedReference");

         Assert.Equal("object", property["Type"]);
         Assert.Equal("CategoryDto", property["Object type"]);
         // The wrapper is what carries nullable, the referenced schema is not nullable at all.
         Assert.Equal("Yes", property["Nullable"]);
      }

      [Fact]
      public void Wrapped_enum_shows_its_type_name_values_and_example()
      {
         var property = PropertyRow("wrappedEnum");

         // Used to read "object" with no values at all, the wrapper has neither type nor enum.
         Assert.Equal("string", property["Type"]);
         Assert.Equal("StatusDto", property["Object type"]);
         Assert.Equal("available, pending, sold", property["Enum"]);
         Assert.Equal("available", property["Example"]);
         Assert.Equal("Yes", property["Nullable"]);
      }

      [Fact]
      public void Wrapper_wins_over_the_referenced_schema()
      {
         var property = PropertyRow("wrappedEnumWithOwnExample");

         Assert.Equal("StatusDto", property["Object type"]);
         Assert.Equal("sold", property["Example"]);
      }

      [Fact]
      public void Format_and_deprecation_of_the_referenced_schema_are_documented()
      {
         var property = PropertyRow("wrappedDeprecated");

         Assert.Equal("string", property["Type"]);
         Assert.Equal("uuid", property["Format"]);
         Assert.Equal("Yes", property["Deprecated"]);
      }

      [Fact]
      public void Composition_of_several_schemas_is_named_after_all_of_them()
      {
         var property = PropertyRow("composition");

         // Two references composed into a type the specification never names. No single name
         // describes it, and the names it is written as do, all of them together.
         Assert.Equal("object", property["Type"]);
         Assert.Equal("All of (CategoryDto, StatusDto)", property["Object type"]);
      }

      [Fact]
      public void Array_of_wrapped_references_is_named_after_the_item_type()
      {
         var property = PropertyRow("wrappedList");

         Assert.Equal("array", property["Type"]);
         Assert.Equal("CategoryDto", property["Object type"]);
      }

      [Fact]
      public void Type_name_does_not_spill_over_the_column_next_to_it()
      {
         // Excel spills the content of a cell over the cells to its right for as long as they hold
         // nothing at all, and a fully qualified type name is far wider than its column. The cell
         // next to a filled Object type has to hold something, an empty string when the schema
         // declares no format.
         var header = _worksheet.Rows()
            .First(row => row.Cell(1).GetString() == "Name")
            .CellsUsed()
            .ToDictionary(cell => cell.GetString(), cell => cell.Address.ColumnNumber);
         var objectTypeColumn = header["Object type"];

         var spilling = _worksheet.RowsUsed()
            .Where(row => !string.IsNullOrEmpty(row.Cell(objectTypeColumn).GetString()))
            .Where(row => row.Cell(objectTypeColumn + 1).Value.IsBlank)
            .Select(row => $"{row.Cell(1).GetString()} (row {row.RowNumber()})")
            .ToList();

         Assert.Equal(header["Format"], objectTypeColumn + 1);
         Assert.Empty(spilling);
      }

      /// <summary>
      /// Every attribute of a property row, keyed by the header of its table. Column positions
      /// depend on the depth of the property tree, the headers do not.
      /// </summary>
      private Dictionary<string, string> PropertyRow(string name)
      {
         var nameCell = _worksheet.CellsUsed(cell => cell.GetString() == name).Single();
         var headerRow = Enumerable.Range(1, nameCell.Address.RowNumber - 1)
            .Reverse()
            .First(row => _worksheet.Cell(row, 1).GetString() == "Name");

         return _worksheet.Row(headerRow).CellsUsed()
            .Where(header => header.Address.ColumnNumber > 1)
            .ToDictionary(
               header => header.GetString(),
               header => _worksheet.Cell(nameCell.Address.RowNumber, header.Address.ColumnNumber).GetString());
      }

      public sealed class GeneratedWorkbook : IDisposable
      {
         private const string Document = """
            openapi: 3.0.1
            info:
              title: Wrapped references
              version: v1
            paths:
              /pets:
                post:
                  operationId: AddPet
                  requestBody:
                    content:
                      application/json:
                        schema: { $ref: '#/components/schemas/PetDto' }
                  responses:
                    '200': { description: OK }
            components:
              schemas:
                PetDto:
                  type: object
                  properties:
                    directReference: { $ref: '#/components/schemas/CategoryDto' }
                    wrappedReference:
                      allOf: [ { $ref: '#/components/schemas/CategoryDto' } ]
                      nullable: true
                    wrappedEnum:
                      allOf: [ { $ref: '#/components/schemas/StatusDto' } ]
                      nullable: true
                    wrappedEnumWithOwnExample:
                      allOf: [ { $ref: '#/components/schemas/StatusDto' } ]
                      example: sold
                    wrappedDeprecated:
                      allOf: [ { $ref: '#/components/schemas/DeprecatedDto' } ]
                    composition:
                      allOf:
                        - { $ref: '#/components/schemas/CategoryDto' }
                        - { $ref: '#/components/schemas/StatusDto' }
                    wrappedList:
                      type: array
                      items:
                        allOf: [ { $ref: '#/components/schemas/CategoryDto' } ]
                        nullable: true
                CategoryDto:
                  type: object
                  properties:
                    name: { type: string }
                StatusDto:
                  type: string
                  enum: [ available, pending, sold ]
                  example: available
                DeprecatedDto:
                  type: string
                  format: uuid
                  deprecated: true
            """;

         private readonly string _outputFile =
            Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");

         public GeneratedWorkbook()
         {
            using var input = new MemoryStream(Encoding.UTF8.GetBytes(Document));
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
