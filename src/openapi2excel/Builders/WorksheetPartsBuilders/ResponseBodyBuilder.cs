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
            AddPropertiesTreeForMediaTypes(response.Value.Content, attributesColumnIndex);
        }
        AddEmptyRow();
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


}