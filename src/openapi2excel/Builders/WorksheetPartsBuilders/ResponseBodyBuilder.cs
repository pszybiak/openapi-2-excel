using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

internal class ResponseBodyBuilder(
    RowPointer actualRow,
    int attributesColumnIndex,
    IXLWorksheet worksheet,
    OpenApiDocumentationOptions options) : WorksheetPartBuilder(actualRow, worksheet, options)
{
    public void AddResponseBodyPart(OpenApiOperation operation)
    {
        if (!operation.Responses.Any())
            return;

        foreach (var response in operation.Responses)
        {
            AddResponseHttpCode(response.Key, response.Value.Description);
            foreach (var mediaType in response.Value.Content)
            {
                AddResponseFormat(mediaType.Key);
                AddResponseHeader();
                foreach (var property in mediaType.Value.Schema.Properties)
                {
                    AddProperty(property.Key, property.Value, 1, attributesColumnIndex);
                }
                AddEmptyRow();
            }
            AddEmptyRow();
        }
        AddEmptyRow();
    }

    private void AddResponseHeader()
    {
        var lastUsedColumn = FillSchemaDescriptionHeaderCells(attributesColumnIndex);
        ActualRow.MovePrev();
        FillHeaderBackground(1, lastUsedColumn);
        ActualRow.MoveNext(2);
    }

    private void AddResponseHttpCode(string httpCode, string? description)
    {
        if (string.IsNullOrEmpty(description))
        {
            Fill(1).WithText($"Response HttpCode: {httpCode}").WithBoldStyle();
        }
        else
        {
            Fill(1).WithText($"Response HttpCode: {httpCode}: {description}").WithBoldStyle();
        }
        ActualRow.MoveNext();
    }

    private void AddResponseFormat(string name)
    {
        Fill(1).WithText($"Response format: {name}").WithBoldStyle();
        ActualRow.MoveNext();
    }
}