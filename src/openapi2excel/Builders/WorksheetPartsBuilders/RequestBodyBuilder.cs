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

        AddRequestBodyHeader();
        foreach (var mediaType in operation.RequestBody.Content)
        {
            AddContentTypeRow(mediaType.Key);
            foreach (var property in mediaType.Value.Schema.Properties)
            {
                AddRequestParameter(property.Key, property.Value, 1);
            }
            AddEmptyRow();
        }
        AddEmptyRow();
    }

    private void AddRequestBodyHeader()
    {
        FillHeaderCell("Name", 1);
        for (var columnIndex = 2; columnIndex < attributesColumnIndex; columnIndex++)
        {
            FillHeaderCell(null, columnIndex);
        }
        FillSchemaDescriptionHeaderCells(attributesColumnIndex);
        MoveToNextRow();
    }

    private void AddContentTypeRow(string name)
    {
        FillCell(1, "Request format");
        FillCell(attributesColumnIndex, name);
        for (var i = 1; i < attributesColumnIndex + 3; i++) // TODO: fill background of all attributes cells
        {
            FillCell(i, XLColor.LightBlue);
        }
        MoveToNextRow();
    }

    private void AddRequestParameter(string name, OpenApiSchema schema, int level)
    {
        AddPropertyRow(name, schema, level++);
        if (schema.Items != null)
        {
            AddPropertyRow("element", schema.Items, level);
        }

        foreach (var property in schema.Properties)
        {
            AddRequestParameter(property.Key, property.Value, level);
        }
    }

    private void AddPropertyRow(string name, OpenApiSchema schema, int level)
    {
        for (var columnIndex = 1; columnIndex < level; columnIndex++)
        {
            FillCell(columnIndex, XLColor.DarkGray);
        }
        FillCell(level, name);
        FillSchemaDescriptionCells(schema, attributesColumnIndex);
        MoveToNextRow();
    }
}