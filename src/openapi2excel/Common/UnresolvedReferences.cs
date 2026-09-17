using Microsoft.OpenApi.Models;

namespace openapi2excel.core.Common;

/// <summary>
/// References the reader of the specification leaves for the document to resolve. Nothing here
/// repairs a violation of the standard, so none of it is a sanitization the tool reports or the
/// options switch off.
/// </summary>
internal static class UnresolvedReferences
{
   /// <summary>
   /// Replaces the schema of a response header written as a <c>$ref</c> with the schema the
   /// components section declares.
   /// <para>
   /// The reader resolves a reference everywhere else, so a header stating one reaches the document
   /// as a name and nothing else, and is documented as a bare object instead of as the type it is.
   /// </para>
   /// </summary>
   public static void ResolveResponseHeaderSchemas(OpenApiDocument document)
   {
      var declarations = document.Components?.Schemas;
      if (declarations is null || declarations.Count == 0 || document.Paths is null)
      {
         return;
      }

      foreach (var response in document.Paths.Values
                  .SelectMany(pathItem => pathItem.Operations?.Values ?? Enumerable.Empty<OpenApiOperation>())
                  .SelectMany(operation => operation.Responses?.Values ?? Enumerable.Empty<OpenApiResponse>()))
      {
         foreach (var header in response.Headers?.Values ?? Enumerable.Empty<OpenApiHeader>())
         {
            if (Declared(declarations, header.Schema) is { } declared)
            {
               header.Schema = declared;
            }
         }
      }
   }

   /// <summary>
   /// The schema a reference points at, when the reference is the only thing the schema carries.
   /// Null when there is nothing to resolve, or when the document declares no such schema.
   /// </summary>
   private static OpenApiSchema? Declared(IDictionary<string, OpenApiSchema> declarations, OpenApiSchema? schema)
   {
      // A reference to another document names a schema this one never read. A schema of the same
      // name declared here is a different schema, and documenting it as that one would invent
      // everything it says.
      if (schema?.Reference is not { IsExternal: false, Id.Length: > 0 }
          || schema.Reference.Id is not { Length: > 0 } name
          || schema.Type is not null
          || schema.Properties.Any()
          || schema.Items is not null)
      {
         return null;
      }

      return declarations.TryGetValue(name, out var declared) && !ReferenceEquals(declared, schema)
         ? declared
         : null;
   }
}
