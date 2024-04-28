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
      Worksheet.Cell(ActualRow, startColumn).SetBackground(propertyLevel - 1, HeaderBackgroundColor);

      var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
      schemaDescriptor.AddNameValue(propertyName, ActualRow, propertyLevel);
      schemaDescriptor.AddSchemaDescriptionValues(propertySchema, ActualRow, attributesColumnIndex);
      ActualRow.MoveNext();
   }

   protected int AddSchemaDescriptionHeader()
   {
      const int startColumn = 1;

      var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
      schemaDescriptor.AddNameHeader(ActualRow, startColumn);
      var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, attributesColumnIndex);

      Worksheet.Cell(ActualRow, startColumn)
         .SetBackground(lastUsedColumn, HeaderBackgroundColor)
         .SetBottomBorder(lastUsedColumn);

      ActualRow.MoveNext();
      return lastUsedColumn;
   }
}