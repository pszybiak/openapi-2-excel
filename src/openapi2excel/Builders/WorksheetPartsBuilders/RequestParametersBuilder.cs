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

        AddRequestParametersHeader();
        operation.Parameters.ForEach(AddPropertyRow);
        AddEmptyRow();
    }

    private void AddRequestParametersHeader()
    {
        FillHeader(1).WithText("Parameters");
        ActualRow.MoveNext();

        var nextCell = FillHeader(1).WithText("Name")
            .Next(attributesColumnIndex - 1).WithText("Location")
            .Next().WithText("Serialization")
            .Next().WithText("Required")
            .Next().GetCell();
        var lastUsedColumn = FillSchemaDescriptionHeaderCells(nextCell.Address.ColumnNumber);
        ActualRow.MovePrev();
        FillHeaderBackground(1, lastUsedColumn);
        ActualRow.MoveNext();
        FillHeaderBackground(1, lastUsedColumn);
        ActualRow.MoveNext();
    }

    private void AddPropertyRow(OpenApiParameter parameter)
    {
        var currentColumn = attributesColumnIndex;
        FillCell(1, parameter.Name);
        FillCell(currentColumn++, parameter.In.ToString()?.ToUpper());
        FillCell(currentColumn++, parameter.Style?.ToString());
        FillCell(currentColumn++, parameter.Required);
        FillSchemaDescriptionCells(parameter.Schema, currentColumn);
        ActualRow.MoveNext();
    }
}