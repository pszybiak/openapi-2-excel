using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using OpenApi2Excel.Common;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders;

internal class RequestParametersBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddRequestParametersPart(OpenApiOperation operation)
   {
      if (!operation.Parameters.Any())
         return;

      Fill(1).WithText("PARAMETERS").WithBoldStyle();
      ActualRow.MoveNext();
      using (var _ = new Section(Worksheet, ActualRow))
      {
         var nextCell = Fill(1).WithText("Name").WithBoldStyle()
            .Next(attributesColumnIndex - 1).WithText("Location").WithBoldStyle()
            .Next().WithText("Serialization").WithBoldStyle()
            .Next().WithText("Required").WithBoldStyle()
            .Next().GetCell();

         FillSchemaDescriptionHeaderCells(nextCell.Address.ColumnNumber);
         ActualRow.MoveNext();
         operation.Parameters.ForEach(AddPropertyRow);
         ActualRow.MovePrev();
      }

      ActualRow.MoveNext(2);
   }

   private void AddPropertyRow(OpenApiParameter parameter)
   {
      var nextCell = Fill(1).WithText(parameter.Name)
         .Next(attributesColumnIndex - 1).WithText(parameter.In.ToString()?.ToUpper())
         .Next().WithText(parameter.Style?.ToString())
         .Next().WithText(Options.Language.Get(parameter.Required))
         .Next().GetCell();

      FillSchemaDescriptionCells(parameter.Schema, nextCell.Address.ColumnNumber);
      ActualRow.MoveNext();
   }
}