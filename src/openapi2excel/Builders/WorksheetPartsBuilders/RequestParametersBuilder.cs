using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using OpenApi2Excel.Common;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

internal class RequestParametersBuilder(RowPointer actualRow, int attributesColumnIndex, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
    : WorksheetPartBuilder(actualRow, worksheet, options)
{
    public void AddRequestParametersPart(OpenApiOperation operation)
    {
        if (!operation.Parameters.Any())
            return;

        AddRequestParametersHeader(operation);
        operation.Parameters.ForEach(AddPropertyRow);
        AddEmptyRow();
    }

    private void AddRequestParametersHeader(OpenApiOperation operation)
    {
        var columnIndex = attributesColumnIndex;
        FillHeaderCell("Parameters", 1);
        MoveToNextRow();
        FillHeaderCell("Name", 1);
        FillHeaderCell("Location", columnIndex++);
        FillHeaderCell("Serialization", columnIndex++);
        FillHeaderCell("Required", columnIndex++);
        var lastUsedColumn = FillSchemaDescriptionHeaderCells(columnIndex);
        FillHeaderBackground(1, lastUsedColumn);
        MoveToNextRow();
    }

    private void AddPropertyRow(OpenApiParameter parameter)
    {
        var currentColumn = attributesColumnIndex;
        FillCell(1, parameter.Name);
        FillCell(currentColumn++, parameter.In.ToString()?.ToUpper());
        FillCell(currentColumn++, parameter.Style?.ToString());
        FillCell(currentColumn++, parameter.Required);
        FillSchemaDescriptionCells(parameter.Schema, currentColumn);
        MoveToNextRow();
    }
}