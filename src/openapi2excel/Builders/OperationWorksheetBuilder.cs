using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace openapi2excel.Builders;

internal class OperationWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
    : WorksheetBuilder(options)
{
    private int _actualRowIndex = 1;
    private IXLWorksheet _worksheet = null!;
    private int _attributesColumnsStartIndex;

    public IXLWorksheet Build(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
    {
        _actualRowIndex = 1;
        _worksheet = workbook.Worksheets.Add(operation.OperationId);

        SetMaxTreeLevel(operation);
        for (var columnIndex = 1; columnIndex <= _attributesColumnsStartIndex; columnIndex++)
            _worksheet.Column(columnIndex).Width = 3;

        AddHomePageLink();
        AddEmptyRow();
        AddOperationInfos(path, pathItem, operationType, operation);
        AddEmptyRow();
        AddRequestParameters(operation);
        AddRequestBody(operation);

        return _worksheet;
    }

    private void SetMaxTreeLevel(OpenApiOperation operation)
    {
        if (operation.RequestBody is null)
            return;

        _attributesColumnsStartIndex = operation.RequestBody.Content
            .Select(openApiMediaType => openApiMediaType.Value.Schema)
            .Select(schema => SetMaxTreeLevel(schema, 0)).Prepend(1).Max();
    }

    private static int SetMaxTreeLevel(OpenApiSchema schema, int currentLevel)
    {
        currentLevel++;
        return schema.Properties == null || !schema.Properties.Any()
            ? currentLevel
            : schema.Properties.Select(schemaProperty => SetMaxTreeLevel(schemaProperty.Value, currentLevel)).Prepend(currentLevel).Max();
    }

    private void AddOperationInfos(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
    {
        AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationType), operationType.ToString().ToUpper());

        AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.Path), path);
        AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.PathDescription), pathItem.Summary);
        AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.PathSummary), pathItem.Summary);

        AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationDescription), operation.Description);
        AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationSummary), operation.Summary);

        AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.Deprecated), Options.Language.Get(operation.Deprecated));
    }

    private void AddRequestParameters(OpenApiOperation operation)
    {
        if (!operation.Parameters.Any())
            return;

        AddInfoRow("Parameters", null, true);
        foreach (var parameter in operation.Parameters)
        {
            AddInfoRow(parameter.Name, parameter.In?.ToString().ToUpper());
        }
        AddEmptyRow();
    }

    private void AddRequestBody(OpenApiOperation operation)
    {
        if (operation.RequestBody is null)
            return;

        foreach (var mediaType in operation.RequestBody.Content)
        {
            AddInfoRow("Request", mediaType.Key, true);
            foreach (var property in mediaType.Value.Schema.Properties)
            {
                AddRequestParameter(property.Key, property.Value, 1);
            }
            AddEmptyRow();
        }

        AddEmptyRow();
    }


    private void AddRequestParameter(string name, OpenApiSchema schema, int level)
    {
        AddPropertyRow(name, schema, level++);
        foreach (var property in schema.Properties)
        {
            AddRequestParameter(property.Key, property.Value, level);
        }
    }

    private void AddHomePageLink()
    {
        _worksheet.Cell(_actualRowIndex, 1).Value = "<<<<<";
        _worksheet.Cell(_actualRowIndex++, 1).SetHyperlink(new XLHyperlink($"'{InfoWorksheetBuilder.Name}'!A1"));
    }

    private void AddEmptyRow()
    {
        _actualRowIndex++;
    }

    private void AddInfoRow(string label, string? value, bool addIfNotExists = false)
    {
        if (!addIfNotExists && value is null)
            return;

        _worksheet.Cell(_actualRowIndex, 1).Value = label;
        _worksheet.Cell(_actualRowIndex++, _attributesColumnsStartIndex + 2).Value = value;
    }

    private void AddPropertyRow(string name, OpenApiSchema schema, int level)
    {
        var column = _attributesColumnsStartIndex + 1;

        _worksheet.Cell(_actualRowIndex, level).Value = name;
        _worksheet.Cell(_actualRowIndex, column++).Value = schema.Type;
        _worksheet.Cell(_actualRowIndex, column++).Value = schema.Format;
        _worksheet.Cell(_actualRowIndex++, column).Value = schema.Description;
    }
}