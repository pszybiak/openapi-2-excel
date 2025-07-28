using Microsoft.OpenApi.Models;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

internal static class MaxPropertiesTreeLevel
{
   public static int Calculate(OpenApiOperation operation, int maxTreeLevel)
   {
      var level = 1;
      if (operation.RequestBody is null && !operation.Responses.Any())
         return level;

      if (operation.RequestBody is not null)
      {
         level = EstablishMaxTreeLevel(operation.RequestBody.Content, maxTreeLevel);
      }

      return operation.Responses.Select(response => EstablishMaxTreeLevel(response.Value.Content, maxTreeLevel))
         .Prepend(level)
         .Max();
   }

   private static int EstablishMaxTreeLevel(IDictionary<string, OpenApiMediaType> mediaTypes, int maxTreeLevel)
      => mediaTypes.Select(openApiMediaType => openApiMediaType.Value.Schema)
         .Select(schema => EstablishMaxTreeLevel(schema, 1, maxTreeLevel))
         .Prepend(1)
         .Max();

   private static int EstablishMaxTreeLevel(OpenApiSchema? schema, int currentLevel, int maxTreeLevel)
   {
      var max = currentLevel;

      if (currentLevel >= maxTreeLevel || schema == null)
      {
         return max;
      }

      if (schema.Items != null)
      {
         max = Math.Max(currentLevel, EstablishMaxTreeLevel(schema.Items, currentLevel + 1, maxTreeLevel));
      }

      // When the schema composition is a conjunction, we merge all the subschemas on the same level
      max = !schema.AllOf.Any()
         ? max
         : schema.AllOf
            .Select(subschema => EstablishMaxTreeLevel(subschema, currentLevel, maxTreeLevel))
            .Prepend(max)
            .Max();

      // When the schema composition is a disjunction and there's
      // many subschemas, they go one level bellow the current level.
      int subschemaLevel;

      subschemaLevel = currentLevel + (schema.AnyOf.Count > 1 ? 1 : 0);
      max = !schema.AnyOf.Any()
         ? max
         : schema.AnyOf
            .Select(subschema => EstablishMaxTreeLevel(subschema, subschemaLevel, maxTreeLevel))
            .Prepend(max)
            .Max();

      subschemaLevel = currentLevel + (schema.OneOf.Count > 1 ? 1 : 0);
      max = !schema.OneOf.Any()
         ? max
         : schema.OneOf
            .Select(subschema => EstablishMaxTreeLevel(subschema, subschemaLevel, maxTreeLevel))
            .Prepend(max)
            .Max();

      max = schema.Properties?.Any() != true
         ? max
         : schema.Properties
            .Select(schemaProperty => EstablishMaxTreeLevel(schemaProperty.Value, currentLevel + 1, maxTreeLevel))
            .Prepend(max)
            .Max();

      return max;
   }
}