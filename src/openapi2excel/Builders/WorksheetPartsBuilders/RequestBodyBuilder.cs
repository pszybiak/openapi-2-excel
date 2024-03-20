﻿using ClosedXML.Excel;
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
            AddContentTypeRow(mediaType.Key);
            AddRequestBodyHeader();
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
        var lastUsedColumn = FillSchemaDescriptionHeaderCells(attributesColumnIndex);
        FillHeaderBackground(1, lastUsedColumn);
        AddBottomBorder(1, lastUsedColumn);
        ActualRow.MoveNext();
    }

    private void AddContentTypeRow(string name)
    {
        FillCell(1, $"Request format: {name}");
        FillHeaderBackground(1, attributesColumnIndex + 2);
        ActualRow.MoveNext();
    }

    private void AddRequestParameter(string name, OpenApiSchema schema, int level)
    {
        // current property
        AddPropertyRow(name, schema, level++);
        if (schema.Items != null)
        {
            if (schema.Items.Properties.Any())
            {
                // array object subproperties
                foreach (var property in schema.Items.Properties)
                {
                    AddRequestParameter(property.Key, property.Value, level);
                }
            }
            else
            {
                // if array contains simple type
                AddRequestParameter("element", schema.Items, level);
            }
        }

        // subproperties
        foreach (var property in schema.Properties)
        {
            AddRequestParameter(property.Key, property.Value, level);
        }
    }

    private void AddPropertyRow(string name, OpenApiSchema schema, int level)
    {
        FillHeaderBackground(1, level - 1);
        FillCell(level, name);
        FillSchemaDescriptionCells(schema, attributesColumnIndex);
        ActualRow.MoveNext();
    }
}