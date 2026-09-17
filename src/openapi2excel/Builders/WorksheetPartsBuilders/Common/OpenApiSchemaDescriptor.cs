using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

internal class OpenApiSchemaDescriptor(
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options,
   ObjectLinkRegistry? objectLinks = null)
{
   public int AddNameHeader(RowPointer actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetTextBold(options.Translation.FieldName).GetColumnNumber();

   public int AddNameValue(string name, int actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetText(name).GetColumnNumber();

   public int AddSchemaDescriptionHeader(RowPointer actualRow, int startColumn)
   {
      var cell = worksheet.Cell(actualRow, startColumn).SetTextBold(options.Translation.FieldType)
         .CellRight().SetTextBold(options.Translation.FieldObjectType)
         .CellRight().SetTextBold(options.Translation.FieldFormat)
         .CellRight().SetTextBold(options.Translation.FieldLength)
         .CellRight().SetTextBold(options.Translation.FieldRequired)
         .CellRight().SetTextBold(options.Translation.FieldNullable)
         .CellRight().SetTextBold(options.Translation.FieldRange)
         .CellRight().SetTextBold(options.Translation.FieldPattern)
         .CellRight().SetTextBold(options.Translation.FieldEnum)
         .CellRight().SetTextBold(options.Translation.FieldDeprecated)
         .CellRight().SetTextBold(options.Translation.FieldDefault)
         .CellRight().SetTextBold(options.Translation.FieldExample)
         .CellRight().SetTextBold(options.Translation.FieldDescription);

      return cell.GetColumnNumber();
   }

   public int AddSchemaDescriptionValues(OpenApiSchema schema, bool required, RowPointer actualRow, int startColumn, string? description = null, bool includeArrayItemType = false)
   {
      // The property attributes live either on the property schema or, when that schema only wraps
      // a reference to add a keyword to it, on the schema it wraps.
      var effective = schema.GetEffectiveSchema();

      if (schema.Items != null && includeArrayItemType)
      {
         var items = schema.Items.GetEffectiveSchema();
         var typeCell = worksheet.Cell(actualRow, startColumn).SetText(options.Translation.ArrayType);
         var cell = AddObjectTypeValue(schema, typeCell)
            .CellRight().SetText(items.Type)
            .CellRight().SetText(FirstNotEmpty(schema.Items.Format, items.Format))
            .CellRight().SetText(schema.GetPropertyLengthDescription(options.Translation)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(required.Translate(options.Translation)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(IsNullable(schema, effective).Translate(options.Translation)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(schema.GetPropertyRangeDescription(options.Translation)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(FirstNotEmpty(schema.Items.Pattern, items.Pattern))
            .CellRight().SetText(FirstNotEmpty(schema.Items.GetEnumDescription(options.Translation), items.GetEnumDescription(options.Translation)))
            .CellRight().SetText(IsDeprecated(schema, effective).Translate(options.Translation)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(FirstNotEmpty(schema.GetExampleDescription(options.Translation), effective.GetExampleDescription(options.Translation)))
            .CellRight().SetText((string.IsNullOrEmpty(schema.Description) ? description : schema.Description).StripHtmlTags());

         return cell.GetColumnNumber();
      }
      else
      {
         var typeCell = worksheet.Cell(actualRow, startColumn).SetText(schema.GetTypeDescription(options.Translation));
         var cell = AddObjectTypeValue(schema, typeCell)
            .CellRight().SetText(FirstNotEmpty(schema.Format, effective.Format))
            .CellRight().SetText(FirstNotEmpty(schema.GetPropertyLengthDescription(options.Translation), effective.GetPropertyLengthDescription(options.Translation))).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(required.Translate(options.Translation)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(IsNullable(schema, effective).Translate(options.Translation)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(FirstNotEmpty(schema.GetPropertyRangeDescription(options.Translation), effective.GetPropertyRangeDescription(options.Translation))).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(FirstNotEmpty(schema.Pattern, effective.Pattern))
            .CellRight().SetText(FirstNotEmpty(schema.GetEnumDescription(options.Translation), effective.GetEnumDescription(options.Translation)))
            .CellRight().SetText(IsDeprecated(schema, effective).Translate(options.Translation)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center)
            .CellRight().SetText(FirstNotEmpty(schema.GetDefaultDescription(options.Translation), effective.GetDefaultDescription(options.Translation)))
            .CellRight().SetText(FirstNotEmpty(schema.GetExampleDescription(options.Translation), effective.GetExampleDescription(options.Translation)))
            .CellRight().SetText((string.IsNullOrEmpty(schema.Description) ? description : schema.Description).StripHtmlTags());

         return cell.GetColumnNumber();
      }
   }

   /// <summary>
   /// Writes the Object type column, right of the type of the property, and remembers the cell when
   /// it names a schema, so that it can be linked to the section documenting that schema.
   /// </summary>
   private IXLCell AddObjectTypeValue(OpenApiSchema schema, IXLCell typeCell)
   {
      var cell = typeCell.CellRight().SetText(schema.GetObjectDescription(options.Translation));
      if (schema.GetSchemaName() is { Length: > 0 } name)
      {
         objectLinks?.Register(cell, name);
      }

      return cell;
   }

   /// <summary>
   /// The wrapper wins, it is the more specific one, and it is what the property declares.
   /// <para>
   /// Never null. Excel spills the content of a cell over the cells to its right for as long as
   /// they hold nothing at all, and a fully qualified type name in Object type spills over the
   /// empty Format next to it. An empty string is not an empty cell, so it stops the spill without
   /// putting anything into the document.
   /// </para>
   /// </summary>
   private static string FirstNotEmpty(string? value, string? fallback)
      => (string.IsNullOrEmpty(value) ? fallback : value) ?? string.Empty;

   private static bool IsNullable(OpenApiSchema schema, OpenApiSchema effective)
      => schema.Nullable || effective.Nullable;

   private static bool IsDeprecated(OpenApiSchema schema, OpenApiSchema effective)
      => schema.Deprecated || effective.Deprecated;
}
