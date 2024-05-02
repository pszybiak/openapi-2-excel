using Microsoft.OpenApi.Models;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

internal class MaxPropertiesTreeLevel
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
      if (schema.Items != null)
      {
         return Math.Max(currentLevel, EstablishMaxTreeLevel(schema.Items, currentLevel + 1));
      }

      return schema.Properties?.Any() != true
         ? currentLevel
         : schema.Properties
            .Select(schemaProperty => EstablishMaxTreeLevel(schemaProperty.Value, currentLevel + 1))
            .Prepend(currentLevel)
            .Max();
   }
}