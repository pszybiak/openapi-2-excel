using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class PropertiesTreeBuilder(
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
{
   protected OpenApiDocumentationOptions Options { get; } = options;
   protected IXLWorksheet Worksheet { get; } = worksheet;
   private RowPointer ActualRow { get; set; } = null!;
   protected XLColor HeaderBackgroundColor => XLColor.LightGray;

   public int AddPropertiesTree(RowPointer actualRow, OpenApiSchema schema)
   {
      ActualRow = actualRow;
      var columnCount = AddSchemaDescriptionHeader();
      var startColumn = CorrectRootElementIfArray(schema) ? 2 : 1;
      AddProperties(schema, startColumn);
      return columnCount;
   }

   protected bool CorrectRootElementIfArray(OpenApiSchema schema)
   {
      if (schema.Items == null)
         return false;

      AddPropertyRow("<array>", schema, 1);
      return true;
   }

   protected void AddProperties(OpenApiSchema schema, int level)
   {
      if (schema.Items != null)
      {
         if (schema.Items.Properties.Any())
         {
            // array of object properties
            foreach (var property in schema.Items.Properties)
            {
               AddProperty(property.Key, property.Value, level);
            }
         }
         else
         {
            // if array contains simple type items
            AddProperty("<value>", schema.Items, level);
         }
      }

      // subproperties
      foreach (var property in schema.Properties)
      {
         AddProperty(property.Key, property.Value, level);
      }
   }

   protected void AddProperty(string name, OpenApiSchema schema, int level)
   {
      AddPropertyRow(name, schema, level++);
      AddProperties(schema, level);
   }

   private void AddPropertyRow(string propertyName, OpenApiSchema propertySchema, int propertyLevel)
   {
      const int startColumn = 1;
      Fill(startColumn).WithBackground(HeaderBackgroundColor, propertyLevel - 1);
      Fill(propertyLevel).WithText(propertyName);
      FillSchemaDescriptionCells(propertySchema, attributesColumnIndex);
      ActualRow.MoveNext();
   }

   protected int AddSchemaDescriptionHeader()
   {
      const int startColumn = 1;
      var cellBuilder = Fill(startColumn).WithText("Name").WithBoldStyle()
         .Next(attributesColumnIndex - 1).WithText("Type").WithBoldStyle()
         .Next().WithText("Format").WithBoldStyle()
         .Next().WithText("Length").WithBoldStyle()
         .Next().WithText("Range").WithBoldStyle()
         .Next().WithText("Pattern").WithBoldStyle()
         .Next().WithText("Deprecated").WithBoldStyle()
         .Next().WithText("Description").WithBoldStyle();

      var lastUsedColumn = cellBuilder.GetCellNumber();

      Fill(startColumn).WithBackground(HeaderBackgroundColor, lastUsedColumn)
         .GoTo(startColumn).WithBottomBorder(lastUsedColumn);

      ActualRow.MoveNext();
      return lastUsedColumn;
   }

   protected void FillSchemaDescriptionCells(OpenApiSchema schema, int startColumn)
   {
      Fill(startColumn).WithText(schema.GetTypeDescription())
         .Next().WithText(schema.Format)
         .Next().WithText(schema.GetPropertyLengthDescription()).WithHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         .Next().WithText(schema.GetPropertyRangeDescription()).WithHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         .Next().WithText(schema.Pattern)
         .Next().WithText(Options.Language.Get(schema.Deprecated)).WithHorizontalAlignment(XLAlignmentHorizontalValues.Center)
         .Next().WithText(schema.GetPropertyDescription());
   }

   protected CellBuilder Fill(int column)
      => new(Worksheet.Cell(ActualRow, column), Options);
}