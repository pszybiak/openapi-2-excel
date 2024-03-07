using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace openapi2excel.Builders;

internal class OperationWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
    : WorksheetBuilder(options)
{
    private int _actualRowIndex = 1;
    private IXLWorksheet _worksheet = null!;

    public IXLWorksheet Build(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
    {
        _actualRowIndex = 1;
        _worksheet = workbook.Worksheets.Add(operation.OperationId);
        _worksheet.Column(1).Width = 20;

        AddHomePageLink();
        AddEmptyRow();
        AddOperationInfos(path, pathItem, operationType, operation);
        AddEmptyRow();

        return _worksheet;
    }

    private void AddOperationInfos(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
    {
        AddRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationType), operationType.ToString().ToUpper());

        AddRow(Options.Language.Get(OpenApiDocumentationLanguageConst.Path), path);
        AddRow(Options.Language.Get(OpenApiDocumentationLanguageConst.PathDescription), pathItem.Summary);
        AddRow(Options.Language.Get(OpenApiDocumentationLanguageConst.PathSummary), pathItem.Summary);

        AddRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationDescription), operation.Description);
        AddRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationSummary), operation.Summary);

        AddRow(Options.Language.Get(OpenApiDocumentationLanguageConst.Deprecated), Options.Language.Get(operation.Deprecated));
    }

    private void AddHomePageLink()
    {
        _worksheet.Cell(_actualRowIndex, 1).Value = "<<<<<";
        _worksheet.Cell(_actualRowIndex++, 1).SetHyperlink(new XLHyperlink($"'{InfoWorksheetBuilder.Name}'!A1"));
    }

    private void AddEmptyRow()
    {
        _actualRowIndex++;
    }

    private void AddRow(string label, string? value, bool addIfNotExists = false)
    {
        if (!addIfNotExists && value is null)
            return;

        _worksheet.Cell(_actualRowIndex, 1).Value = label;
        _worksheet.Cell(_actualRowIndex++, 2).Value = value;
    }
}