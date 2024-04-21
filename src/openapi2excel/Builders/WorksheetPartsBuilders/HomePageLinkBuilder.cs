using ClosedXML.Excel;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class HomePageLinkBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddHomePageLinkPart()
   {
      var cell = Worksheet.Cell(ActualRow, 1);
      cell.SetValue("<<<<<");
      cell.SetHyperlink(new XLHyperlink($"'{InfoWorksheetBuilder.Name}'!A1"));
      ActualRow.MoveNext(2);
   }
}