using System.Text;
using ClosedXML.Excel;
using openapi2excel.core;
using openapi2excel.core.Lang;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// The index of the info worksheet, as it lands in the file: the groups the specification itself
   /// defines, the outline that collapses them, and the links into the operation worksheets.
   /// </summary>
   public class InfoWorksheetIndexTest
   {
      private const string Document = """
         openapi: 3.0.1
         info:
           title: Measurements
           version: v1
         tags:
           - name: device
             description: Everything about devices
           - name: measurement
             description: Readings taken by a device
         paths:
           /measurements:
             summary: The readings themselves
             get:
               operationId: GetMeasurements
               summary: Lists the measurements
               tags: [ measurement ]
               responses:
                 '200':
                   description: Ok
             post:
               operationId: AddMeasurement
               summary: Adds a measurement
               tags: [ measurement, device ]
               responses:
                 '200':
                   description: Ok
           /devices:
             get:
               operationId: GetDevices
               summary: Lists the devices
               tags: [ device ]
               responses:
                 '200':
                   description: Ok
           /health:
             get:
               operationId: GetHealth
               description: |
                 Tells whether the service answers
                 and nothing else
               responses:
                 '200':
                   description: Ok
         """;

      [Fact]
      public void Index_groups_the_operations_by_tag_and_by_path_the_way_the_specification_does()
      {
         using var workbook = Generate("en");
         var info = workbook.Worksheet("Info");

         string[][] expected =
         [
            ["", "", "Version", "v1"],
            ["", "", "Title", "Measurements"],
            ["", "", "", ""],
            ["OPERATIONS", "", "", ""],
            ["Tag", "Path", "Operation type: Id", "Operation summary"],
            // Declared first, and the tag of an operation carrying two of them groups it twice.
            ["device", "", "", "Everything about devices"],
            ["", "/measurements", "", "The readings themselves"],
            ["", "", "POST: AddMeasurement", "Adds a measurement"],
            ["", "/devices", "", ""],
            ["", "", "GET: GetDevices", "Lists the devices"],
            ["measurement", "", "", "Readings taken by a device"],
            ["", "/measurements", "", "The readings themselves"],
            ["", "", "GET: GetMeasurements", "Lists the measurements"],
            ["", "", "POST: AddMeasurement", "Adds a measurement"],
            // An operation carrying no tag has no group of the specification to land in.
            ["Other operations", "", "", ""],
            ["", "/health", "", ""],
            ["", "", "GET: GetHealth", "Tells whether the service answers"]
         ];

         Assert.Equal(expected, Rows(info));
      }

      [Fact]
      public void Every_group_collapses_from_the_header_row_above_it()
      {
         using var workbook = Generate("en");
         var info = workbook.Worksheet("Info");

         // Tag rows stay visible, path rows collapse into them, operation rows into the path.
         Assert.Equal(XLOutlineSummaryVLocation.Top, info.Outline.SummaryVLocation);
         Assert.Equal([0, 0, 0, 0, 0, 0, 1, 2, 1, 2, 0, 1, 2, 2, 0, 1, 2],
            Enumerable.Range(1, 17).Select(row => info.Row(row).OutlineLevel));
      }

      [Fact]
      public void Operation_row_links_into_the_worksheet_documenting_it()
      {
         using var workbook = Generate("en");
         var info = workbook.Worksheet("Info");

         // The operation and its summary both walk to its worksheet.
         Assert.Equal("'AddMeasurement'!A1", info.Cell("C8").GetHyperlink().InternalAddress);
         Assert.Equal("'AddMeasurement'!A1", info.Cell("D8").GetHyperlink().InternalAddress);

         // One link per listed operation, and the operation of two tags is listed under each.
         var links = info.CellsUsed(cell => cell.HasHyperlink && cell.Address.ColumnNumber == 3)
            .Select(cell => cell.GetHyperlink().InternalAddress)
            .ToList();
         Assert.Equal(
            ["'AddMeasurement'!A1", "'GetDevices'!A1", "'GetMeasurements'!A1", "'AddMeasurement'!A1", "'GetHealth'!A1"],
            links);

         // Neither a tag nor a path header links anywhere, they cover several worksheets.
         Assert.False(info.Cell("A6").HasHyperlink);
         Assert.False(info.Cell("B7").HasHyperlink);
      }

      [Fact]
      public void Index_headers_follow_the_language_and_the_names_of_the_specification_do_not()
      {
         using var workbook = Generate("pl");
         var info = workbook.Worksheet("Informacje");

         Assert.Equal("OPERACJE", info.Cell("A4").GetString());
         Assert.Equal(["Tag", "Ścieżka", "Typ operacji: Identyfikator", "Podsumowanie operacji"],
            Enumerable.Range(1, 4).Select(column => info.Cell(5, column).GetString()));
         Assert.Equal("device", info.Cell("A6").GetString());
         Assert.Equal("Pozostałe operacje", info.Cell("A15").GetString());
      }

      [Fact]
      public void Header_of_the_index_is_told_apart_from_the_rows_below_it()
      {
         using var workbook = Generate("en");
         var info = workbook.Worksheet("Info");

         // The labels of the info block name their value the way the index names its columns.
         Assert.True(info.Cell("C1").Style.Font.Bold);
         Assert.False(info.Cell("D1").Style.Font.Bold);
         Assert.True(info.Cell("A4").Style.Font.Bold);
         Assert.True(info.Cell("A5").Style.Font.Bold);
         Assert.Equal(XLColor.LightGray, info.Cell("D5").Style.Fill.BackgroundColor);
         Assert.Equal(XLBorderStyleValues.Medium, info.Cell("D5").Style.Border.BottomBorder);

         // A tag row is filled across the whole index, a path row is only bold.
         var unpainted = info.Cell("A1").Style.Fill.BackgroundColor;
         Assert.True(info.Cell("A6").Style.Font.Bold);
         Assert.Equal(info.Cell("A6").Style.Fill.BackgroundColor, info.Cell("D6").Style.Fill.BackgroundColor);
         Assert.NotEqual(unpainted, info.Cell("D6").Style.Fill.BackgroundColor);
         Assert.NotEqual(XLColor.LightGray, info.Cell("D6").Style.Fill.BackgroundColor);
         Assert.True(info.Cell("B7").Style.Font.Bold);
         Assert.Equal(unpainted, info.Cell("D7").Style.Fill.BackgroundColor);
      }

      [Fact]
      public void Tag_and_path_columns_are_narrow_enough_to_read_as_the_indentation_of_a_group()
      {
         using var workbook = Generate("en");
         var info = workbook.Worksheet("Info");

         // Excel measures a column in characters of the default font of the workbook, Calibri 11,
         // whose widest digit takes 7 pixels, plus 5 pixels of padding: 30 and 40 pixels.
         Assert.Equal(30, info.Column(1).Width * 7 + 5, 1);
         Assert.Equal(40, info.Column(2).Width * 7 + 5, 1);
      }

      [Fact]
      public void Operation_stating_no_id_is_named_by_its_type_alone()
      {
         const string noId = """
            openapi: 3.0.1
            info:
              title: Nameless
              version: v1
            paths:
              /health:
                get:
                  summary: Tells whether the service answers
                  responses:
                    '200':
                      description: Ok
            """;

         using var workbook = Generate("en", noId);
         var info = workbook.Worksheet("Info");

         // The specification names an operation with its operationId, and states none here, so the
         // cell has nothing to state but the type. The worksheet documenting the operation is named
         // after its method and its path instead.
         Assert.Equal(["", "", "GET", "Tells whether the service answers"],
            Enumerable.Range(1, 4).Select(column => info.Cell(8, column).GetString()));
         Assert.Equal("'GET_health'!A1", info.Cell("C8").GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Specification_with_no_operation_at_all_has_no_index()
      {
         const string empty = """
            openapi: 3.0.1
            info:
              title: Nothing
              version: v1
            paths: {}
            """;

         using var workbook = Generate("en", empty);
         var info = workbook.Worksheet("Info");

         Assert.Equal(2, info.LastRowUsed()!.RowNumber());
         Assert.Equal("Title", info.Cell("C2").GetString());
      }

      private static List<string[]> Rows(IXLWorksheet worksheet)
         => Enumerable.Range(1, worksheet.LastRowUsed()!.RowNumber())
            .Select(row => Enumerable.Range(1, 4).Select(column => worksheet.Cell(row, column).GetString()).ToArray())
            .ToList();

      private static XLWorkbook Generate(string language, string document = Document)
      {
         var outputFile = Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");
         try
         {
            using var input = new MemoryStream(Encoding.UTF8.GetBytes(document));
            OpenApiDocumentationGenerator.GenerateDocumentation(input, outputFile,
                  new OpenApiDocumentationOptions { Translation = TranslationsSelector.Load(language).Translation })
               .GetAwaiter().GetResult();

            using var generated = File.OpenRead(outputFile);
            return new XLWorkbook(generated);
         }
         finally
         {
            File.Delete(outputFile);
         }
      }
   }
}
