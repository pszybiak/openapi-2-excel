using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using System.Text;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

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
        FillHeaderCell("Length", column++);
        FillHeaderCell("Range", column++);
        FillHeaderCell("Pattern", column++);
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

        var column = attributesColumnIndex;
        FillCell(level, name);
        FillCell(column++, schema.Type);
        FillCell(column++, schema.Format);
        FillCell(column++, GetPropertyLength(schema));
        FillCell(column++, GetPropertyRange(schema));
        FillCell(column, schema.Pattern);
        FillCell(column, schema.Description);
        MoveToNextRow();
    }

    private static string GetPropertyLength(OpenApiSchema schema)
    {
        StringBuilder propertyTypeDescription = new();
        if (schema.MinLength is not null)
        {
            propertyTypeDescription.Append($"{schema.MinLength}");
        }
        if (schema.MinLength is not null && !(schema.MinLength is null && schema.MaxLength is not null))
        {
            propertyTypeDescription.Append("..");
        }
        if (schema.MaxLength is not null)
        {
            propertyTypeDescription.Append(schema.MaxLength);
        }
        return propertyTypeDescription.ToString();
    }

    private static string GetPropertyRange(OpenApiSchema schema)
    {
        StringBuilder propertyTypeDescription = new();
        if (schema.Minimum is not null)
        {
            var sign = schema.ExclusiveMinimum is null or false ? "[" : "(";
            propertyTypeDescription.Append($"{sign}{schema.Minimum}");
        }
        else if (schema.Maximum is not null)
        {
            propertyTypeDescription.Append("(..");
        }
        if (schema.Minimum is not null || schema.Maximum is not null)
        {
            propertyTypeDescription.Append(';');
        }
        if (schema.Maximum is not null)
        {
            var sign = schema.ExclusiveMaximum is null or false ? "]" : ")";
            propertyTypeDescription.Append($"{schema.Maximum}{sign}");
        }
        else if (schema.Minimum is not null)
        {
            propertyTypeDescription.Append("..)");
        }
        return propertyTypeDescription.ToString();
    }
}