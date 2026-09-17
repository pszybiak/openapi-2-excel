using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class OperationInfoBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddOperationInfoSection(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation)
   {
      Cell(1).SetTextBold(Options.Translation.OperationHeader);
      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var cell = Cell(1).SetTextBold(Options.Translation.OperationType).CellRight(attributesColumnIndex).SetText(operationType.ToString().ToUpper())
            .IfNotEmpty(operation.OperationId, c => c.NextRow().SetTextBold(Options.Translation.OperationId).CellRight(attributesColumnIndex).SetText(operation.OperationId))
            .NextRow().SetTextBold(Options.Translation.OperationPath).CellRight(attributesColumnIndex).SetText(path)
            .IfNotEmpty(pathItem.Description, c => c.NextRow().SetTextBold(Options.Translation.OperationPathDescription).CellRight(attributesColumnIndex).SetText(pathItem.Description))
            .IfNotEmpty(pathItem.Summary, c => c.NextRow().SetTextBold(Options.Translation.OperationPathSummary).CellRight(attributesColumnIndex).SetText(pathItem.Summary))
            .IfNotEmpty(operation.Description, c => c.NextRow().SetTextBold(Options.Translation.OperationDescription).CellRight(attributesColumnIndex).SetText(operation.Description))
            .IfNotEmpty(operation.Summary, c => c.NextRow().SetTextBold(Options.Translation.OperationSummary).CellRight(attributesColumnIndex).SetText(operation.Summary))
            .NextRow().SetTextBold(Options.Translation.OperationDeprecated).CellRight(attributesColumnIndex).SetText(operation.Deprecated.Translate(Options.Translation));

         ActualRow.GoTo(cell.Address.RowNumber);
      }

      ActualRow.MoveNext(2);
   }
}