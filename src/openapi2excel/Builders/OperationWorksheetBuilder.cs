using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace openapi2excel.Builders;

internal class OperationWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
    : WorksheetBuilder(options)
{
    public IXLWorksheet Build(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
    {
        var worksheet = workbook.Worksheets.Add(operation.OperationId);

        var column = 1;
        var row = 1;
        worksheet.Cell(row, column).Value = "<<<<<";
        worksheet.Cell(row++, column).SetHyperlink(new XLHyperlink($"'{InfoWorksheetBuilder.Name}'!A1"));
        worksheet.Cell(row++, column).Value = path;
        worksheet.Cell(row++, column).Value = operationType.ToString();
        worksheet.Cell(row++, column).Value = operation.Summary;

        return worksheet;
    }
}