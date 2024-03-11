using ClosedXML.Excel;

namespace openapi2excel.Builders.WorksheetPartsBuilders;

internal class HomePageLinkBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
    : WorksheetPartBuilder(actualRow, worksheet, options)
{
    public void AddHomePageLinkPart()
    {
        FillCell(1, "<<<<<");
        Worksheet.Cell(ActualRow, 1).SetHyperlink(new XLHyperlink($"'{InfoWorksheetBuilder.Name}'!A1"));
        ActualRow.MoveNext(2);
    }
}