using Microsoft.OpenApi.Models;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

internal static class MaxPropertiesTreeLevel
{
   public static int Calculate(OpenApiOperation operation)
   {
      var level = 1;
      if (operation.RequestBody is null && !operation.Responses.Any())
         return level;

      if (operation.RequestBody is not null)
      {
         level = EstablishMaxTreeLevel(operation.RequestBody.Content);
      }

      return operation.Responses.Select(response => EstablishMaxTreeLevel(response.Value.Content))
         .Prepend(level)
         .Max();
   }

   private static int EstablishMaxTreeLevel(IDictionary<string, OpenApiMediaType> mediaTypes)
      => mediaTypes.Select(openApiMediaType => openApiMediaType.Value.Schema)
         .Select(schema => EstablishMaxTreeLevel(schema, 1))
         .Prepend(1)
         .Max();

   private static int EstablishMaxTreeLevel(OpenApiSchema schema, int currentLevel)
   {
      var max = currentLevel;
      if (schema.Items != null)
      {
         max = Math.Max(currentLevel, EstablishMaxTreeLevel(schema.Items, currentLevel + 1));
      }

      max = !schema.AllOf.Any()
         ? max
         : schema.AllOf
            .Select(schemaProperty => EstablishMaxTreeLevel(schemaProperty, currentLevel))
            .Prepend(max)
            .Max();

      max = !schema.AnyOf.Any()
         ? max
         : schema.AnyOf
            .Select(schemaProperty => EstablishMaxTreeLevel(schemaProperty, currentLevel))
            .Prepend(max)
            .Max();

      max = schema.Properties?.Any() != true
         ? max
         : schema.Properties
            .Select(schemaProperty => EstablishMaxTreeLevel(schemaProperty.Value, currentLevel + 1))
            .Prepend(max)
            .Max();

      return max;
   }
}