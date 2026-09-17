using System.Text;
using ClosedXML.Excel;
using openapi2excel.core;
using openapi2excel.core.Lang;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// The document itself, generated in another language: labels, worksheet names, the links that
   /// point at them, and the culture used for the values taken from the specification.
   /// </summary>
   public class LocalizedDocumentTest
   {
      private const string Document = """
         openapi: 3.0.1
         info:
           title: Measurements
           version: v1
         paths:
           /measurements:
             post:
               operationId: AddMeasurement
               requestBody:
                 content:
                   application/json:
                     schema:
                       type: object
                       required: [ weight ]
                       properties:
                         weight:
                           type: number
                           format: double
                           minimum: 0.5
                           maximum: 99.75
                           default: 2.25
                           example: 1.5
                         name:
                           type: string
                           minLength: 2
                           maxLength: 40
               responses:
                 '200':
                   description: Created
         """;

      [Fact]
      public void Worksheet_names_labels_and_links_follow_the_language()
      {
         using var polish = Generate("pl");
         var info = polish.Worksheet("Informacje");
         var operation = polish.Worksheet("AddMeasurement");

         Assert.Equal("Wersja", info.Cell("C1").GetString());
         Assert.Equal("Tytuł", info.Cell("C2").GetString());
         Assert.Equal("INFORMACJE O OPERACJI", operation.Cell("A3").GetString());
         Assert.Equal("Typ operacji", operation.Cell("A4").GetString());

         // The link back to the info worksheet has to follow its translated name.
         Assert.Equal("'Informacje'!A1", operation.Cell("A1").GetHyperlink().InternalAddress);

         var operationLink = info.CellsUsed(cell => cell.HasHyperlink).First();
         Assert.Equal("POST: AddMeasurement", operationLink.GetString());
         Assert.Equal("'AddMeasurement'!A1", operationLink.GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Composed_labels_keep_what_the_placeholder_stands_for()
      {
         using var polish = Generate("pl");
         var texts = polish.Worksheet("AddMeasurement").CellsUsed().Select(cell => cell.GetString()).ToList();

         Assert.Contains("Format treści: application/json", texts);
         Assert.Contains("Kod HTTP odpowiedzi: 200: Created", texts);
         Assert.Contains("ŻĄDANIE", texts);
         Assert.Contains("ODPOWIEDŹ", texts);
      }

      [Fact]
      public void Yes_and_no_follow_the_language()
      {
         using var polish = Generate("pl");
         var operation = polish.Worksheet("AddMeasurement");
         var weight = Row(operation, "weight");

         Assert.Contains("Tak", weight);
         Assert.Contains("Nie", weight);
         Assert.DoesNotContain("Yes", weight);
         Assert.DoesNotContain("No", weight);
      }

      [Fact]
      public void Numbers_are_formatted_with_the_culture_of_the_translation()
      {
         using var english = Generate("en");
         using var polish = Generate("pl");

         // en-US and pl-PL disagree about the decimal separator, and the document has to be the
         // same on every machine, whatever culture that machine runs.
         Assert.Contains("1.5", Row(english.Worksheet("AddMeasurement"), "weight"));
         Assert.Contains("2.25", Row(english.Worksheet("AddMeasurement"), "weight"));
         Assert.Contains("[0.5;99.75]", Row(english.Worksheet("AddMeasurement"), "weight"));

         Assert.Contains("1,5", Row(polish.Worksheet("AddMeasurement"), "weight"));
         Assert.Contains("2,25", Row(polish.Worksheet("AddMeasurement"), "weight"));
         Assert.Contains("[0,5;99,75]", Row(polish.Worksheet("AddMeasurement"), "weight"));
      }

      [Fact]
      public void Lengths_do_not_depend_on_the_language()
      {
         using var english = Generate("en");
         using var polish = Generate("pl");

         Assert.Contains("2..40", Row(english.Worksheet("AddMeasurement"), "name"));
         Assert.Contains("2..40", Row(polish.Worksheet("AddMeasurement"), "name"));
      }

      [Fact]
      public void Translated_document_has_content_wherever_the_english_one_has()
      {
         // A label left null by a translation would silently blank a cell, so the two documents have
         // to fill exactly the same cells.
         using var english = Generate("en");
         using var polish = Generate("pl");

         Assert.Equal(english.Worksheets.Count, polish.Worksheets.Count);

         var blanks = new List<string>();
         foreach (var (englishSheet, polishSheet) in english.Worksheets.Zip(polish.Worksheets))
         {
            foreach (var cell in englishSheet.CellsUsed(cell => !string.IsNullOrEmpty(cell.GetString())))
            {
               var translated = polishSheet.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber).GetString();
               if (string.IsNullOrEmpty(translated))
               {
                  blanks.Add($"{englishSheet.Name}!{cell.Address} ('{cell.GetString()}')");
               }
            }
         }

         Assert.Empty(blanks);
      }

      private static List<string> Row(IXLWorksheet worksheet, string name)
      {
         var nameCell = worksheet.CellsUsed(cell => cell.GetString() == name).Single();
         return worksheet.Row(nameCell.Address.RowNumber).CellsUsed()
            .Select(cell => cell.GetString())
            .Where(text => !string.IsNullOrEmpty(text))
            .ToList();
      }

      private static XLWorkbook Generate(string language)
      {
         var outputFile = Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");
         try
         {
            using var input = new MemoryStream(Encoding.UTF8.GetBytes(Document));
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
