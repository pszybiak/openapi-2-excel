using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

internal class OpenApiSchemaDescriptor(IXLWorksheet worksheet, OpenApiDocumentationOptions options)
{
   public int AddNameHeader(RowPointer actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetTextBold("Name").GetColumnNumber();

   public int AddNameValue(string name, int actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetText(name).GetColumnNumber();

   public int AddSchemaDescriptionHeader(RowPointer actualRow, int startColumn)
   {
      var cell = worksheet.Cell(actualRow, startColumn).SetTextBold("Type")
         .CellRight().SetTextBold("Object type")
         .CellRight().SetTextBold("Format")
         .CellRight().SetTextBold("Length")
         .CellRight().SetTextBold("Required")
         .CellRight().SetTextBold("Nullable")
         .CellRight().SetTextBold("Range")
         .CellRight().SetTextBold("Pattern")
         .CellRight().SetTextBold("Enum")
         .CellRight().SetTextBold("Deprecated")
         .CellRight().SetTextBold("Default")
         .CellRight().SetTextBold("Example")
         .CellRight().SetTextBold("Description");

      return cell.GetColumnNumber();
   }

   public int AddSchemaDescriptionValues(OpenApiSchema schema, bool required, RowPointer actualRow, int startColumn, string? description = null, bool includeArrayItemType = false )
   {
      if (schema.Items != null && includeArrayItemType)
      {
         var cell = worksheet.Cell(actualRow, startColumn).SetText("array")
            .CellRight().SetText(schema.GetObjectDescription())
            .CellRight().SetText(schema.Items.Type)
            .CellRight().SetText(schema.Items.Format)
            .CellRight().SetText(schema.GetPropertyLengthDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(options.Language.Get(required)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(options.Language.Get(schema.Nullable)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(schema.GetPropertyRangeDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(schema.Items.Pattern)
            .CellRight().SetText(schema.Items.GetEnumDescription())
            .CellRight().SetText(options.Language.Get(schema.Deprecated)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(schema.GetExampleDescription())
            .CellRight().SetText((string.IsNullOrEmpty(schema.Description) ? description : schema.Description).StripHtmlTags());

         return cell.GetColumnNumber();
      }
      else
      {
         var cell = worksheet.Cell(actualRow, startColumn).SetText(schema.GetTypeDescription())
            .CellRight().SetText(schema.GetObjectDescription())
            .CellRight().SetText(schema.Format)
            .CellRight().SetText(schema.GetPropertyLengthDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(options.Language.Get(required)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(options.Language.Get(schema.Nullable)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(schema.GetPropertyRangeDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(schema.Pattern)
            .CellRight().SetText(schema.GetEnumDescription())
            .CellRight().SetText(options.Language.Get(schema.Deprecated)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(schema.GetDefaultDescription())
            .CellRight().SetText(schema.GetExampleDescription())
            .CellRight().SetText((string.IsNullOrEmpty(schema.Description) ? description : schema.Description).StripHtmlTags());

         return cell.GetColumnNumber();
      }
   }
}