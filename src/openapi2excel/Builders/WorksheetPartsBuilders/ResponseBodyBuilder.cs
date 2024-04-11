using ClosedXML.Excel;
using Microsoft.OpenApi.Models;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class ResponseBodyBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options) : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddResponseBodyPart(OpenApiOperation operation)
   {
      if (!operation.Responses.Any())
         return;

      Fill(1).WithText("RESPONSE").WithBoldStyle();
      AddEmptyRow();
      using (var _ = new Section(Worksheet, ActualRow))
      {
         foreach (var response in operation.Responses)
         {
            AddResponseHttpCode(response.Key, response.Value.Description);
            AddReponseHeaders(response.Value.Headers);
            AddPropertiesTreeForMediaTypes(response.Value.Content, attributesColumnIndex);
         }
      }

      AddEmptyRow();
   }

   private void AddReponseHeaders(IDictionary<string, OpenApiHeader> valueHeaders)
   {
      if (!valueHeaders.Any())
         return;

      AddEmptyRow();

      Fill(1).WithText("Response headers").WithBoldStyle();
      ActualRow.MoveNext();

      var nextCell = Fill(1).WithText("Name").WithBoldStyle()
         .Next(attributesColumnIndex - 1).WithText("Required").WithBoldStyle()
         .Next().WithText("Deprecated").WithBoldStyle()
         .Next().GetCellNumber();

      int lastUsedColumn = FillSchemaDescriptionHeaderCells(nextCell);
      ActualRow.MovePrev();
      FillHeaderBackground(1, lastUsedColumn);
      ActualRow.MoveNext(2);

      foreach (var openApiHeader in valueHeaders)
      {
         var nextCellNumber = Fill(1).WithText(openApiHeader.Key)
            .Next(attributesColumnIndex - 1).WithText(Options.Language.Get(openApiHeader.Value.Required))
            .Next().WithText(Options.Language.Get(openApiHeader.Value.Deprecated))
            .Next().GetCellNumber();
         nextCellNumber = FillSchemaDescriptionCells(openApiHeader.Value.Schema, nextCellNumber);
         Fill(nextCellNumber).WithText(openApiHeader.Value.Description);
         ActualRow.MoveNext();
      }

      ActualRow.MoveNext();
   }

   private void AddResponseHttpCode(string httpCode, string? description)
   {
      if (string.IsNullOrEmpty(description))
      {
         Fill(1).WithText($"Response HttpCode: {httpCode}").WithBoldStyle();
      }
      else
      {
         Fill(1).WithText($"Response HttpCode: {httpCode}: {description}").WithBoldStyle();
      }

      ActualRow.MoveNext();
   }
}