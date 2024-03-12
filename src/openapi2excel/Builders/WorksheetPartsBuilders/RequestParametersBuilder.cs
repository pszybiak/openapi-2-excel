using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

internal class RequestParametersBuilder(RowPointer actualRow, int attributesColumnIndex, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
    : WorksheetPartBuilder(actualRow, worksheet, options)
{
    public void AddRequestParametersPart(OpenApiOperation operation)
    {
        if (!operation.Parameters.Any())
            return;

        AddRequestParametersHeader(operation);
        foreach (var parameter in operation.Parameters)
        {
            AddPropertyRow(parameter.Name, parameter);
        }
        AddEmptyRow();
    }

    private void AddRequestParametersHeader(OpenApiOperation operation)
    {
        FillCell(1, "Parameters");
    }

    private void AddPropertyRow(string name, OpenApiParameter parameter)
    {
        var column = attributesColumnIndex;
        FillCell(1, name);
        FillCell(column, parameter.In.ToString()?.ToUpper());
        // TODO: more parameters
        MoveToNextRow();
    }
}