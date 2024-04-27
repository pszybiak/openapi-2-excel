using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders
{
   internal abstract class WorksheetPartBuilder(
      RowPointer actualRow,
      IXLWorksheet worksheet,
      OpenApiDocumentationOptions options)
   {
      protected OpenApiDocumentationOptions Options { get; } = options;
      protected RowPointer ActualRow { get; } = actualRow;
      protected IXLWorksheet Worksheet { get; } = worksheet;
      protected XLColor HeaderBackgroundColor => XLColor.LightGray;

      protected void AddEmptyRow()
         => ActualRow.MoveNext();

      protected IXLCell Cell(int column)
         => Worksheet.Cell(ActualRow.Get(), column);

      protected CellBuilder Fill(int column)
         => new(Worksheet.Cell(ActualRow, column), Options);

      protected CellBuilder FillHeader(int column)
         => Fill(column).WithBackground(HeaderBackgroundColor);

      protected void AddPropertiesTreeForMediaTypes(IDictionary<string, OpenApiMediaType> mediaTypes,
         int attributesColumnIndex)
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
         foreach (var mediaType in mediaTypes)
         {
            var bodyFormatRowPointer = ActualRow.Copy();
            Cell(1).SetTextBold($"Body format: {mediaType.Key}");
            ActualRow.MoveNext();

            var columnCount = builder.AddPropertiesTree(ActualRow, mediaType.Value.Schema);
            Worksheet.Cell(bodyFormatRowPointer, 1).SetBackground(columnCount, HeaderBackgroundColor);
            AddEmptyRow();
         }
      }

      protected void AddProperty(string name, OpenApiSchema schema, int level, int attributesColumnIndex)
      {
         // current property
         AddPropertyRow(name, schema, level++);
         AddProperties(schema, level, attributesColumnIndex);

         // ReSharper disable once SeparateLocalFunctionsWithJumpStatement
         void AddPropertyRow(string propertyName, OpenApiSchema propertySchema, int propertyLevel)
         {
            FillHeaderBackground(1, propertyLevel - 1);
            FillCell(propertyLevel, propertyName);
            FillSchemaDescriptionCells(propertySchema, attributesColumnIndex);
            ActualRow.MoveNext();
         }
      }

      protected void AddProperties(OpenApiSchema schema, int level, int attributesColumnIndex)
      {
         if (schema.Items != null)
         {
            if (schema.Items.Properties.Any())
            {
               // array of object properties
               foreach (var property in schema.Items.Properties)
               {
                  AddProperty(property.Key, property.Value, level, attributesColumnIndex);
               }
            }
            else
            {
               // if array contains simple type items
               AddProperty("element", schema.Items, level, attributesColumnIndex);
            }
         }

         // subproperties
         foreach (var property in schema.Properties)
         {
            AddProperty(property.Key, property.Value, level, attributesColumnIndex);
         }
      }

      protected int FillSchemaDescriptionCells(OpenApiSchema schema, int startColumn)
      {
         var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
         return schemaDescriptor.AddSchemaDescriptionValues(schema, ActualRow, startColumn);
      }

      protected int FillSchemaDescriptionHeaderCells(int attributesStartColumn)
      {
         var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);
         schemaDescriptor.AddNameHeader(ActualRow, 1);
         var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, attributesStartColumn);

         Fill(1).WithBackground(HeaderBackgroundColor, lastUsedColumn)
            .GoTo(1).WithBottomBorder(lastUsedColumn);

         return lastUsedColumn;
      }

      protected void FillCell(int column, string? value, XLColor? backgoundColor = null,
         XLAlignmentHorizontalValues alignment = XLAlignmentHorizontalValues.Left)
      {
         var cellBuilder = Fill(column).WithText(value);
         if (backgoundColor is not null)
         {
            cellBuilder.WithBackground(backgoundColor);
         }

         cellBuilder.WithHorizontalAlignment(alignment);
      }

      protected void FillBackground(int startColumn, int endColumn, XLColor backgroundColor)
      {
         var cellBuilder = Fill(startColumn);
         for (var columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
         {
            cellBuilder.WithBackground(backgroundColor).Next();
         }
      }

      protected void FillHeaderBackground(int startColumn, int endColumn)
         => FillBackground(startColumn, endColumn, HeaderBackgroundColor);
   }
}