using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace openapi2excel.Builders.WorksheetPartsBuilders;

internal class RequestBodyBuilder(
    RowPointer actualRow,
    int attributesColumnIndex,
    IXLWorksheet worksheet,
    OpenApiDocumentationOptions options)
    : WorksheetPartBuilder(actualRow, worksheet, options)
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
        var column = attributesColumnIndex;
        FillHeaderCell("Field name", 1);
        for (var i = 2; i < column; i++)
        {
            FillHeaderCell(null, i);
        }
        FillHeaderCell("Type", column++);
        FillHeaderCell("Format", column++);
        FillHeaderCell("Description", column);
        MoveToNextRow();
        return;

        void FillHeaderCell(string? label, int columnIndex) => FillCell(columnIndex, label, XLColor.LightBlue);
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
            AddPropertyRow("value", schema.Items, level);
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

        var column = attributesColumnIndex;
        FillCell(level, name);
        FillCell(column++, schema.Type);
        FillCell(column++, schema.Format);
        FillCell(column, schema.Description);
        MoveToNextRow();
    }
}