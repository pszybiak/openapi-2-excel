using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

internal class RequestBodyBuilder(
    RowPointer actualRow,
    int attributesColumnIndex,
    IXLWorksheet worksheet,
    OpenApiDocumentationOptions options) : WorksheetPartBuilder(actualRow, worksheet, options)
{
    public void AddRequestBodyPart(OpenApiOperation operation)
    {
        if (operation.RequestBody is null)
            return;

        foreach (var mediaType in operation.RequestBody.Content)
        {
            AddRequestFormat(mediaType.Key);
            AddRequestHeader();
            foreach (var property in mediaType.Value.Schema.Properties)
            {
                AddProperty(property.Key, property.Value, 1, attributesColumnIndex);
            }
            AddEmptyRow();
        }
        AddEmptyRow();
    }

    private void AddRequestHeader()
    {
        var lastUsedColumn = FillSchemaDescriptionHeaderCells(attributesColumnIndex);
        ActualRow.MovePrev();
        FillHeaderBackground(1, lastUsedColumn);
        ActualRow.MoveNext(2);
    }

    private void AddRequestFormat(string name)
    {
        Fill(1).WithText($"Request format: {name}").WithBoldStyle();
        ActualRow.MoveNext();
    }
}