using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class OperationInfoBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddOperationInfoPart(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation)
   {
      Cell(1).SetTextBold("OPERATION INFORMATION");
      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var cell = Cell(1).SetTextBold("Operation type").CellRight(attributesColumnIndex)
            .SetText(operationType.ToString().ToUpper())
            .NextRow().SetTextBold("Path").CellRight(attributesColumnIndex).SetText(path)
            .NextRow().SetTextBold("Path description").CellRight(attributesColumnIndex).SetText(pathItem.Description)
            .NextRow().SetTextBold("Path summary").CellRight(attributesColumnIndex).SetText(pathItem.Summary)
            .NextRow().SetTextBold("Operation description").CellRight(attributesColumnIndex)
            .SetText(operation.Description)
            .NextRow().SetTextBold("Operation summary").CellRight(attributesColumnIndex).SetText(operation.Summary)
            .NextRow().SetTextBold("Deprecated").CellRight(attributesColumnIndex)
            .SetText(Options.Language.Get(operation.Deprecated));

         ActualRow.GoTo(cell.Address.RowNumber);
      }

      ActualRow.MoveNext(2);
   }
}