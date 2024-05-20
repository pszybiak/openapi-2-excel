using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class PropertiesTreeBuilder(
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
{
   private readonly int _attributesColumnIndex = attributesColumnIndex + 2;
   protected OpenApiDocumentationOptions Options { get; } = options;
   protected IXLWorksheet Worksheet { get; } = worksheet;
   private RowPointer ActualRow { get; set; } = null!;
   protected XLColor HeaderBackgroundColor => XLColor.LightGray;

   public void AddPropertiesTreeForMediaTypes(RowPointer actualRow, IDictionary<string, OpenApiMediaType> mediaTypes)
   {
      ActualRow = actualRow;
      foreach (var mediaType in mediaTypes)
      {
         var bodyFormatRowPointer = ActualRow.Copy();
         Worksheet.Cell(ActualRow, 1).SetTextBold($"Body format: {mediaType.Key}");
         ActualRow.MoveNext();

         using var _ = new Section(Worksheet, ActualRow);
         var columnCount = AddPropertiesTree(ActualRow, mediaType.Value.Schema);
         Worksheet.Cell(bodyFormatRowPointer, 1).SetBackground(columnCount, HeaderBackgroundColor);

         ActualRow.MoveNext();
      }
   }

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

      AddPropertyRow("<array>", schema, false, 1);
      return true;
   }

   protected void AddProperties(OpenApiSchema schema, int level)
   {
      if (schema.Items != null)
      {
         AddPropertiesForArray(schema, level);
      }
      if (schema.AllOf.Count == 1)
      {
         AddProperties(schema.AllOf[0], level);
      }
      if (schema.AnyOf.Count == 1)
      {
         AddProperties(schema.AnyOf[0], level);
      }
      foreach (var property in schema.Properties)
      {
         AddProperty(property.Key, property.Value, schema.Required.Contains(property.Key), level);
      }
   }

   private void AddPropertiesForArray(OpenApiSchema schema, int level)
   {
      if (schema.Items.Properties.Any())
      {
         // array of object properties
         AddProperties(schema.Items, level);
      }
      else
      {
         // if array contains simple type items
         AddProperty("<value>", schema.Items, false, level);
      }
   }

   protected void AddProperty(string name, OpenApiSchema schema, bool required, int level)
   {
      AddPropertyRow(name, schema, required, level++);
      AddProperties(schema, level);
   }

   private void AddPropertyRow(string propertyName, OpenApiSchema propertySchema, bool required, int propertyLevel)
   {
      const int startColumn = 1;
      Worksheet.Cell(ActualRow, startColumn).SetBackground(propertyLevel - 1, HeaderBackgroundColor);

      var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
      schemaDescriptor.AddNameValue(propertyName, ActualRow, propertyLevel);
      schemaDescriptor.AddSchemaDescriptionValues(propertySchema, required, ActualRow, _attributesColumnIndex);
      ActualRow.MoveNext();
   }

   protected int AddSchemaDescriptionHeader()
   {
      const int startColumn = 1;

      var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
      schemaDescriptor.AddNameHeader(ActualRow, startColumn);
      var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, _attributesColumnIndex);

      Worksheet.Cell(ActualRow, startColumn)
         .SetBackground(lastUsedColumn, HeaderBackgroundColor)
         .SetBottomBorder(lastUsedColumn);

      ActualRow.MoveNext();
      return lastUsedColumn;
   }
}