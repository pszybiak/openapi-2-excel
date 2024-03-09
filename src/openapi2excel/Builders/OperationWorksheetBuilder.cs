using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace openapi2excel.Builders;

internal class OperationWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
    : WorksheetBuilder(options)
{
    private int _actualRowIndex = 1;
    private const int ColumnOffset = 4;
    private IXLWorksheet _worksheet = null!;
    private int _attributesColumnsStartIndex;

    public IXLWorksheet Build(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
    {
        _actualRowIndex = 1;
        _worksheet = workbook.Worksheets.Add(operation.OperationId);

        SetMaxTreeLevel(operation);
        for (var columnIndex = 1; columnIndex < _attributesColumnsStartIndex; columnIndex++)
        {
            _worksheet.Column(columnIndex).Width = 2;
        }

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
            .Select(schema => SetMaxTreeLevel(schema, 0)).Prepend(1).Max() + ColumnOffset;
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
        var column = _attributesColumnsStartIndex;
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

    private void AddHomePageLink()
    {
        FillCell(1, "<<<<<");
        _worksheet.Cell(_actualRowIndex++, 1).SetHyperlink(new XLHyperlink($"'{InfoWorksheetBuilder.Name}'!A1"));
    }

    private void AddEmptyRow() => MoveToNextRow();

    private void AddInfoRow(string label, string? value, bool addIfNotExists = false)
    {
        if (!addIfNotExists && value is null)
            return;

        FillCell(1, label);
        FillCell(_attributesColumnsStartIndex, value);
        MoveToNextRow();
    }

    private void AddContentTypeRow(string name)
    {
        FillCell(1, "Request format");
        FillCell(_attributesColumnsStartIndex, name);
        for (var i = 1; i < _attributesColumnsStartIndex + 3; i++) // TODO: fill background of all attributes cells
        {
            FillCell(i, XLColor.LightBlue);
        }
        MoveToNextRow();
    }

    private void AddPropertyRow(string name, OpenApiSchema schema, int level)
    {
        for (var columnIndex = 1; columnIndex < level; columnIndex++)
        {
            FillCell(columnIndex, XLColor.DarkGray);
        }

        var column = _attributesColumnsStartIndex;
        FillCell(level, name);
        FillCell(column++, schema.Type);
        FillCell(column++, schema.Format);
        FillCell(column, schema.Description);
        MoveToNextRow();
    }

    private void FillCell(int column, string? value, XLColor? backgoundColor = null)
    {
        _worksheet.Cell(_actualRowIndex, column).Value = value;
        if (backgoundColor is not null)
        {
            _worksheet.Cell(_actualRowIndex, column).Style.Fill.BackgroundColor = backgoundColor;
        }
    }

    private void FillCell(int column, XLColor backgoundColor)
    {
        _worksheet.Cell(_actualRowIndex, column).Style.Fill.BackgroundColor = backgoundColor;
    }

    private void MoveToNextRow()
    {
        _actualRowIndex++;
    }
}