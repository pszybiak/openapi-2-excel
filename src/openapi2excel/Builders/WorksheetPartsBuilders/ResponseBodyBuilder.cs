using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class ResponseBodyBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options,
   ObjectLinkRegistry? objectLinks = null) : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddResponseBodyPart(OpenApiOperation operation)
   {
      if (!operation.Responses.Any())
         return;

      Cell(1).SetTextBold(Options.Translation.ResponseHeader);
      ActualRow.MoveNext();
      using (var _ = new Section(Worksheet, ActualRow))
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options, objectLinks);
         foreach (var response in operation.Responses)
         {
            AddResponseHttpCode(response.Key, response.Value.Description);
            AddReponseHeaders(response.Value.Headers);
            builder.AddPropertiesTreeForMediaTypes(ActualRow, response.Value.Content, Options);
         }
      }
      ActualRow.MoveNext();
   }

   private void AddReponseHeaders(IDictionary<string, OpenApiHeader> valueHeaders)
   {
      if (!valueHeaders.Any())
         return;

      ActualRow.MoveNext();

      var responseHeadertRowPointer = ActualRow.Copy();
      Cell(1).SetTextBold(Options.Translation.ResponseHeaders);
      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options, objectLinks);

         InsertHeader(schemaDescriptor);
         ActualRow.MoveNext();

         foreach (var openApiHeader in valueHeaders)
         {
            InsertProperty(openApiHeader, schemaDescriptor);
            ActualRow.MoveNext();
         }
      }
      ActualRow.MoveNext();

      void InsertHeader(OpenApiSchemaDescriptor schemaDescriptor)
      {
         var nextCell = Cell(1).SetTextBold(Options.Translation.FieldName)
            .CellRight(attributesColumnIndex + 1).GetColumnNumber();

         var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, nextCell);

         Worksheet.Cell(ActualRow, 1)
            .SetBackground(lastUsedColumn, HeaderBackgroundColor)
            .SetBottomBorder(lastUsedColumn);

         Worksheet.Cell(responseHeadertRowPointer, 1)
            .SetBackground(lastUsedColumn, HeaderBackgroundColor);
      }

      void InsertProperty(KeyValuePair<string, OpenApiHeader> openApiHeader, OpenApiSchemaDescriptor schemaDescriptor)
      {
         var nextCellNumber = Cell(1).SetText(openApiHeader.Key)
            .CellRight(attributesColumnIndex + 1).GetColumnNumber();

         nextCellNumber = schemaDescriptor.AddSchemaDescriptionValues(openApiHeader.Value.Schema, openApiHeader.Value.Required, ActualRow, nextCellNumber);

         Cell(nextCellNumber).SetText(openApiHeader.Value.Description);
      }
   }

   private void AddResponseHttpCode(string httpCode, string? description)
   {
      var responseCode = httpCode.Equals("default")
         ? Options.Translation.ResponseDefault
         : string.Format(Options.Translation.ResponseHttpCode, httpCode);
      if (!string.IsNullOrEmpty(description) && !description.Equals("default response"))
      {
         responseCode = string.Format(Options.Translation.ResponseWithDescription, responseCode, description);
      }

      Cell(1).SetTextBold(responseCode);

      ActualRow.MoveNext();
   }
}