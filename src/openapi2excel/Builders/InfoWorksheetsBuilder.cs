using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace OpenApi2Excel.Builders;

internal class InfoWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
   : WorksheetBuilder(options)
{
   private OpenApiDocument _readResultOpenApiDocument = null!;
   private IXLWorksheet _infoWorksheet = null!;
   public static string Name => OpenApiDocumentationLanguageConst.Info;
   private int _actualRowIndex = 1;

   public IXLWorksheet Build(OpenApiDocument openApiDocument)
   {
      _readResultOpenApiDocument = openApiDocument;
      _infoWorksheet = workbook.Worksheets.Add(Name);
      _infoWorksheet.Column(1).Width = 11;

      AddVersion();
      AddTitle();
      AddDescription();
      AddEmptyRow();

      return _infoWorksheet;
   }

   public void AddLink(OperationType operation, string path, IXLWorksheet worksheet)
   {
      _infoWorksheet.Cell(_actualRowIndex, 1).Value = operation.ToString().ToUpper();
      _infoWorksheet.Cell(_actualRowIndex, 2).Value = path;

      _infoWorksheet.Cell(_actualRowIndex, 1).SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));
      _infoWorksheet.Cell(_actualRowIndex, 2).SetHyperlink(new XLHyperlink($"'{worksheet.Name}'!A1"));
      _actualRowIndex++;
   }

   private void AddVersion()
   {
      if (string.IsNullOrEmpty(_readResultOpenApiDocument.Info.Version))
      {
         return;
      }

      FillInfo(OpenApiDocumentationLanguageConst.Version, _readResultOpenApiDocument.Info.Version);
   }

   private void AddDescription()
   {
      if (string.IsNullOrEmpty(_readResultOpenApiDocument.Info.Description))
      {
         return;
      }

      FillInfo(OpenApiDocumentationLanguageConst.Description, _readResultOpenApiDocument.Info.Description, true);
   }

   private void AddTitle()
   {
      if (string.IsNullOrEmpty(_readResultOpenApiDocument.Info.Title))
      {
         return;
      }

      FillInfo(OpenApiDocumentationLanguageConst.Title, _readResultOpenApiDocument.Info.Title);
   }

   private void AddEmptyRow()
   {
      _actualRowIndex++;
   }

   private void FillInfo(string languageKey, string value, bool multipleRowText = false)
   {
      _infoWorksheet.Cell(_actualRowIndex, 1).Value = Options.Language.Get(languageKey);
      if (multipleRowText)
      {
         var splitValues = value.Split('\n', '\r', StringSplitOptions.RemoveEmptyEntries).Select(v => v.Trim()).Where(v => !string.IsNullOrEmpty(v));
         foreach (var splitValue in splitValues)
         {
            _infoWorksheet.Cell(_actualRowIndex++, 2).Value = splitValue;
         }
      }
      else
      {
         _infoWorksheet.Cell(_actualRowIndex++, 2).Value = value;
      }
   }
}