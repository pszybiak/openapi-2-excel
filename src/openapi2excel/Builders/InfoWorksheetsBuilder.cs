using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders;

internal class InfoWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
   : WorksheetBuilder(options)
{
   private OpenApiDocument _readResultOpenApiDocument = null!;
   private IXLWorksheet _worksheet = null!;
   public static string Name => OpenApiDocumentationLanguageConst.Info;
   private int _actualRowIndex = 1;

   public IXLWorksheet Build(OpenApiDocument openApiDocument)
   {
      _readResultOpenApiDocument = openApiDocument;
      _worksheet = workbook.Worksheets.Add(Name);
      _worksheet.Column(1).Width = 11;

      AddVersion();
      AddTitle();
      AddDescription();

      return _worksheet;
   }

   public void AddLink(OperationType operation, string path, IXLWorksheet worksheet)
   {
      var cell = _worksheet.Cell(_actualRowIndex, 1);
      cell.SetValue(operation.ToString().ToUpper());
      cell.SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));

      cell = cell.CellRight();
      cell.SetValue(path);
      cell.SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));

      _actualRowIndex++;
   }

   private void AddVersion() => FillInfo(OpenApiDocumentationLanguageConst.Version, _readResultOpenApiDocument.Info.Version);

   private void AddDescription() => FillInfo(OpenApiDocumentationLanguageConst.Description, _readResultOpenApiDocument.Info.Description, true);

   private void AddTitle() => FillInfo(OpenApiDocumentationLanguageConst.Title, _readResultOpenApiDocument.Info.Title);

   private void FillInfo(string languageKey, string value, bool splitMultipleRowText = false)
   {
      if (string.IsNullOrEmpty(value))
      {
         return;
      }

      _worksheet.Cell(_actualRowIndex, 1).Value = Options.Language.Get(languageKey);
      if (splitMultipleRowText)
      {
         value.Split('\n', '\r', StringSplitOptions.RemoveEmptyEntries)
            .Select(v => v.Trim())
            .Where(v => !string.IsNullOrEmpty(v))
            .ForEach(splittedValue => _worksheet.Cell(_actualRowIndex++, 2).Value = splittedValue);
      }
      else
      {
         _worksheet.Cell(_actualRowIndex++, 2).Value = value;
      }
   }
}