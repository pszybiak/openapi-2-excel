using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class OpenApiSchemaDescriptor(IXLWorksheet worksheet, OpenApiDocumentationOptions options)
{
   public int AddNameHeader(RowPointer actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetTextBold("Name").GetColumnNumber();

   public int AddNameValue(string name, int actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetText(name).GetColumnNumber();

   public int AddSchemaDescriptionHeader(RowPointer actualRow, int startColumn)
   {
      var cell = worksheet.Cell(actualRow, startColumn).SetTextBold("Type")
         .CellRight().SetTextBold("Format")
         .CellRight().SetTextBold("Length")
         .CellRight().SetTextBold("Nullable")
         .CellRight().SetTextBold("Range")
         .CellRight().SetTextBold("Pattern")
         .CellRight().SetTextBold("Deprecated")
         .CellRight().SetTextBold("Description");

      return cell.GetColumnNumber();
   }

   public int AddSchemaDescriptionValues(OpenApiSchema schema, RowPointer actualRow, int startColumn)
   {
      var cell = worksheet.Cell(actualRow, startColumn).SetText(schema.Type)
         .CellRight().SetText(schema.Format)
         .CellRight().SetText(schema.GetPropertyLengthDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         .CellRight().SetText(options.Language.Get(schema.Nullable)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         .CellRight().SetText(schema.GetPropertyRangeDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         .CellRight().SetText(schema.Pattern)
         .CellRight().SetText(options.Language.Get(schema.Deprecated)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         .CellRight().SetText(schema.Description);

      return cell.GetColumnNumber();
   }
}