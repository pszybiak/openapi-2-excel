using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

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

      Cell(1).SetTextBold("REQUEST");
      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
         builder.AddPropertiesTreeForMediaTypes(ActualRow, operation.RequestBody.Content);
         ActualRow.MovePrev();
      }

      ActualRow.MoveNext(2);
   }
}