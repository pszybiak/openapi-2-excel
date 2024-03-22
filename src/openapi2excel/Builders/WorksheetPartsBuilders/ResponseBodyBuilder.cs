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

        foreach (var reponse in operation.Responses)
        {
            var httpCode = reponse.Key;
            var responseContent = reponse.Value.Content;
            foreach (var mediaType in responseContent)
            {
                AddResponseBodyHeader(mediaType.Key);
                foreach (var property in mediaType.Value.Schema.Properties)
                {
                    AddResponseParameter(property.Key, property.Value, 1);
                }
                AddEmptyRow();
            }
            AddEmptyRow();
        }
        AddEmptyRow();
    }

    private void AddResponseParameter(string propertyKey, OpenApiSchema propertyValue, int level)
    {
    }

    private void AddResponseBodyHeader(string mediaTypeKey)
    {

    }
}