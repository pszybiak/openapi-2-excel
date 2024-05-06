using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

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

      Cell(1).SetTextBold("RESPONSE");
      ActualRow.MoveNext();
      using (var _ = new Section(Worksheet, ActualRow))
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
         foreach (var response in operation.Responses)
         {
            AddResponseHttpCode(response.Key, response.Value.Description);
            AddReponseHeaders(response.Value.Headers);
            builder.AddPropertiesTreeForMediaTypes(ActualRow, response.Value.Content);
         }
      }
      ActualRow.MoveNext();
   }

   private void AddReponseHeaders(IDictionary<string, OpenApiHeader> valueHeaders)
   {
      if (!valueHeaders.Any())
         return;

      ActualRow.MoveNext();

      Cell(1).SetTextBold("Response headers");
      ActualRow.MoveNext();

      var nextCell = Cell(1).SetTextBold("Name")
         .CellRight(attributesColumnIndex - 1).SetTextBold("Required")
         .CellRight().GetColumnNumber();

      var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
      var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, nextCell);

      Worksheet.Cell(ActualRow, 1)
         .SetBackground(lastUsedColumn, HeaderBackgroundColor)
         .SetBottomBorder(lastUsedColumn);

      ActualRow.MoveNext();

      foreach (var openApiHeader in valueHeaders)
      {
         var nextCellNumber = Cell(1).SetText(openApiHeader.Key)
            .CellRight(attributesColumnIndex - 1).SetText(Options.Language.Get(openApiHeader.Value.Required))
            .CellRight().GetColumnNumber();

         nextCellNumber = schemaDescriptor.AddSchemaDescriptionValues(openApiHeader.Value.Schema, ActualRow, nextCellNumber);
         Cell(nextCellNumber).SetText(openApiHeader.Value.Description);

         ActualRow.MoveNext();
      }
      ActualRow.MoveNext();
   }

   private void AddResponseHttpCode(string httpCode, string? description)
   {
      Cell(1).SetTextBold(string.IsNullOrEmpty(description)
         ? $"Response HttpCode: {httpCode}"
         : $"Response HttpCode: {httpCode}: {description}");

      ActualRow.MoveNext();
   }
}