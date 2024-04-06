using ClosedXML.Excel;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

internal class HomePageLinkBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddHomePageLinkPart()
   {
      FillCell(1, "<<<<<");
      Worksheet.Cell(ActualRow, 1).SetHyperlink(CreateHyperlinkToInfoWorksheet());
      ActualRow.MoveNext(2);
   }

   private XLHyperlink CreateHyperlinkToInfoWorksheet() => new($"'{InfoWorksheetBuilder.Name}'!A1");
}