using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders
{
    internal class OperationInfoBuilder(RowPointer actualRow, int attributesColumnIndex, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
        : WorksheetPartBuilder(actualRow, worksheet, options)
    {
        public void AddOperationInfoPart(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
        {
            AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationType), operationType.ToString().ToUpper());

            AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.Path), path);
            AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.PathDescription), pathItem.Summary);
            AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.PathSummary), pathItem.Summary);

            AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationDescription), operation.Description);
            AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.OperationSummary), operation.Summary);

            AddInfoRow(Options.Language.Get(OpenApiDocumentationLanguageConst.Deprecated), Options.Language.Get(operation.Deprecated));
            AddEmptyRow();
        }

        private void AddInfoRow(string label, string? value, bool addIfNotExists = false)
        {
            if (!addIfNotExists && value is null)
                return;

            Fill(1).WithText(label)
                .Next(attributesColumnIndex - 1)
                .WithText(value!);
            MoveToNextRow();
        }
    }
}
