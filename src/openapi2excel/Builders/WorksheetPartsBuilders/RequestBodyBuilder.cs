using ClosedXML.Excel;
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

      Fill(1).WithText("REQUEST").WithBoldStyle();
      ActualRow.MoveNext();
      using (var _ = new Section(Worksheet, ActualRow))
      {
         AddPropertiesTreeForMediaTypes(operation.RequestBody.Content, attributesColumnIndex);
         ActualRow.MovePrev();
      }
      ActualRow.MoveNext(2);
   }
}