using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class PropertiesTreeBuilder(
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options,
   ObjectLinkRegistry? objectLinks = null)
{
   private readonly int _attributesColumnIndex = attributesColumnIndex + 2;
   protected OpenApiDocumentationOptions Options { get; } = options;
   protected IXLWorksheet Worksheet { get; } = worksheet;
   private RowPointer ActualRow { get; set; } = null!;
   protected static XLColor HeaderBackgroundColor => XLColor.LightGray;

   public void AddPropertiesTreeForMediaTypes(RowPointer actualRow, IDictionary<string, OpenApiMediaType> mediaTypes, OpenApiDocumentationOptions options)
   {
      ActualRow = actualRow;
      foreach (var mediaType in mediaTypes)
      {
         var title = string.Format(Options.Translation.BodyFormat, mediaType.Key);
         if (mediaType.Value.Schema is { } schema)
         {
            AddTitledPropertiesTree(ActualRow, title, schema, options);
         }
         else
         {
            // Nothing to describe, the media type is named and that is all the document knows.
            Worksheet.Cell(ActualRow, 1).SetTextBold(title);
            ActualRow.MoveNext(2);
         }
      }
   }

   /// <summary>
   /// Writes one block of the document: a bold title over the table describing a schema, with the
   /// table collapsing into that title. A body format of the request or of the response is such a
   /// block, and so is one object of the objects worksheet.
   /// </summary>
   /// <param name="titleNamesTheSchema">
   /// The title is the name of the schema itself, the way the objects worksheet titles a section,
   /// and not what holds it, the way a body format does.
   /// </param>
   /// <returns>The last column of the table, the one holding the description of a property.</returns>
   public int AddTitledPropertiesTree(RowPointer actualRow, string title, OpenApiSchema schema,
      OpenApiDocumentationOptions options, bool titleNamesTheSchema = false)
   {
      ActualRow = actualRow;
      var titleRowPointer = ActualRow.Copy();

      // A body format titles the block by what carries the schema and not by the schema itself,
      // while the table below only names what the properties of that schema hold. The name of the
      // schema in the title is what tells a reader which object the whole block is, and it links to
      // the section documenting that object.
      var name = titleNamesTheSchema ? null : schema.GetOwnSchemaName();
      Worksheet.Cell(ActualRow, 1).SetTextBold(name is { Length: > 0 }
         ? string.Format(Options.Translation.BodyFormatWithObject, title, name)
         : title);

      if (name is { Length: > 0 })
      {
         // The title spills over the empty cells to its right, and the cell holding it is two
         // characters wide, so a reader clicks the name on cells that hold nothing at all.
         for (var column = 1; column < _attributesColumnIndex; column++)
         {
            objectLinks?.Register(Worksheet.Cell(ActualRow, column), name);
         }
      }

      ActualRow.MoveNext();

      int columnCount;
      using (var _ = new Section(Worksheet, ActualRow))
      {
         columnCount = AddPropertiesTree(ActualRow, schema, options, titleNamesTheSchema);
         Worksheet.Cell(titleRowPointer, 1).SetBackground(columnCount, HeaderBackgroundColor);
         ActualRow.MovePrev();
      }

      ActualRow.MoveNext(2);
      return columnCount;
   }

   public int AddPropertiesTree(RowPointer actualRow, OpenApiSchema schema,
      OpenApiDocumentationOptions options, bool titleNamesTheSchema = false)
   {
      ActualRow = actualRow;
      var columnCount = AddSchemaDescriptionHeader();
      var firstPropertyRow = ActualRow.Get();

      var startColumn = CorrectRootElementIfArray(schema) ? 2 : 1;
      AddProperties(schema, startColumn, options);

      if (ActualRow.Get() == firstPropertyRow && HoldsNothing(schema))
      {
         // A schema holding nothing, an enum or a simple type declared under a name, still has
         // something to say: its own type, values and description, in the one row the table would
         // otherwise not have at all. The row describes the schema the title already names, so on
         // the objects worksheet it names no other object and repeats no description.
         AddPropertyRow(Options.Translation.ValuePlaceholder, schema, false, 1,
            linkObjectType: !titleNamesTheSchema);

         if (titleNamesTheSchema)
         {
            const int descriptionColumnOffset = 12;
            Worksheet.Cell(firstPropertyRow, _attributesColumnIndex + 1).SetText(string.Empty);
            Worksheet.Cell(firstPropertyRow, _attributesColumnIndex + descriptionColumnOffset)
               .SetText(string.Empty);
         }
      }

      return columnCount;
   }

   /// <summary>
   /// The schema opens nothing under it: no property, no item, no schema it is composed of. A
   /// schema whose properties the depth limit cut off is not one of them, and calling it a value
   /// would say the opposite of what the specification says.
   /// </summary>
   private static bool HoldsNothing(OpenApiSchema schema)
   {
      var effective = schema.GetEffectiveSchema();
      return !effective.Properties.Any()
             && effective.Items is null
             && effective.AllOf.Count == 0
             && effective.AnyOf.Count == 0
             && effective.OneOf.Count == 0;
   }

   protected bool CorrectRootElementIfArray(OpenApiSchema schema)
   {
      if (schema.Items == null)
         return false;

      AddPropertyRow(Options.Translation.ArrayPlaceholder, schema, false, 1);
      return true;
   }

   protected void AddProperties(OpenApiSchema? schema, int level, OpenApiDocumentationOptions options)
   {
      if (schema == null)
         return;

      if (schema.Items != null)
      {
         AddPropertiesForArray(schema, level, options);
      }

      // allOf is one schema written in parts, so the parts are documented together, as the one
      // schema they describe.
      foreach (var part in schema.AllOf)
      {
         AddProperties(part, level, options);
      }

      AddAlternatives(schema.AnyOf, level, options);
      AddAlternatives(schema.OneOf, level, options);

      foreach (var property in schema.Properties)
      {
         AddProperty(property.Key, property.Value, schema.Required.Contains(property.Key), level, options);
      }
   }

   /// <summary>
   /// anyOf and oneOf are alternatives, not parts: a single one is the schema itself, and each of
   /// several is documented as a row of its own, the way a property naming a schema is. Writing
   /// their properties together would read as one object holding all of them.
   /// </summary>
   private void AddAlternatives(IList<OpenApiSchema> alternatives, int level,
      OpenApiDocumentationOptions options)
   {
      if (alternatives.Count == 1)
      {
         AddProperties(alternatives[0], level, options);
         return;
      }

      foreach (var alternative in alternatives)
      {
         // An alternative declared in place carries no name to put in the row naming it.
         if (alternative.GetSchemaName() is { Length: > 0 } name)
         {
            AddProperty(name, alternative, false, level, options);
         }
      }
   }

   private void AddPropertiesForArray(OpenApiSchema schema, int level, OpenApiDocumentationOptions options)
   {
      if (schema.Items.Properties.Any())
      {
         // array of object properties
         AddProperties(schema.Items, level, options);
      }
      else
      {
         // if array contains simple type items
         AddProperty(Options.Translation.ValuePlaceholder, schema.Items, false, level, options);
      }
   }

   protected void AddProperty(string name, OpenApiSchema? schema, bool required, int level, OpenApiDocumentationOptions options)
   {
      if (schema == null || level >= options.MaxDepth)
      {
         return;
      }

      AddPropertyRow(name, schema, required, level++);
      AddProperties(schema, level, options);
   }

   private void AddPropertyRow(string propertyName, OpenApiSchema propertySchema, bool required,
      int propertyLevel, bool linkObjectType = true)
   {
      const int startColumn = 1;
      Worksheet.Cell(ActualRow, startColumn).SetBackground(propertyLevel - 1, HeaderBackgroundColor);

      var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options, linkObjectType ? objectLinks : null);
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
