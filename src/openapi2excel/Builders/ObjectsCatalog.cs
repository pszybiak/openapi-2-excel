using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders;

/// <summary>
/// One named schema of the specification, as the objects worksheet documents it.
/// </summary>
internal sealed class DocumentedObject(string name, OpenApiSchema schema)
{
   /// <summary>Name the specification gives the schema, the one the Object type column shows.</summary>
   public string Name { get; } = name;

   /// <summary>The schema itself, with the reference already resolved.</summary>
   public OpenApiSchema Schema { get; } = schema;
}

/// <summary>
/// The named schemas the operations of a document use: every schema this document declares under a
/// name and an operation reaches, directly or through another one. That is a superset of what the
/// Object type column of a worksheet names, so every name the document shows has a section to link
/// to, and a schema a reader can only find by its name is documented as well.
/// </summary>
internal static class ObjectsCatalog
{
   /// <summary>
   /// Collects the objects, by name, sorted so that a reader can look one up. A schema used by
   /// several operations is one object, the way the specification declares it once.
   /// </summary>
   public static IReadOnlyList<DocumentedObject> Collect(OpenApiDocument document)
   {
      var found = new Dictionary<string, OpenApiSchema>(StringComparer.Ordinal);

      // The reader resolves a reference into the schema it points at, so the same schema is the
      // same instance everywhere. Walking an instance once is enough, and a schema referencing
      // itself, directly or through another one, does not loop.
      var visited = new HashSet<OpenApiSchema>(SchemaIdentity.Comparer);

      var declarations = document.Components?.Schemas;
      foreach (var schema in RootSchemas(document))
      {
         Collect(schema, declarations, found, visited);
      }

      return found
         .OrderBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
         .ThenBy(entry => entry.Key, StringComparer.Ordinal)
         .Select(entry => new DocumentedObject(entry.Key, entry.Value))
         .ToList();
   }

   /// <summary>
   /// Every schema an operation worksheet documents: the parameters of the operation, its request
   /// body, and the headers and the body of every response. A parameter declared by the path
   /// instead of by the operation lands in no worksheet, so it names no object either.
   /// </summary>
   private static IEnumerable<OpenApiSchema?> RootSchemas(OpenApiDocument document)
   {
      if (document.Paths is null)
      {
         yield break;
      }

      foreach (var pathItem in document.Paths.Values)
      {
         foreach (var operation in pathItem.Operations?.Values ?? Enumerable.Empty<OpenApiOperation>())
         {
            foreach (var parameter in operation.Parameters ?? Enumerable.Empty<OpenApiParameter>())
            {
               yield return parameter.Schema;
            }

            foreach (var schema in Contents(operation.RequestBody?.Content))
            {
               yield return schema;
            }

            foreach (var response in operation.Responses?.Values ?? Enumerable.Empty<OpenApiResponse>())
            {
               foreach (var header in response.Headers?.Values ?? Enumerable.Empty<OpenApiHeader>())
               {
                  yield return header.Schema;
               }

               foreach (var schema in Contents(response.Content))
               {
                  yield return schema;
               }
            }
         }
      }
   }

   private static IEnumerable<OpenApiSchema?> Contents(IDictionary<string, OpenApiMediaType>? content)
      => content?.Values.Select(mediaType => mediaType.Schema) ?? Enumerable.Empty<OpenApiSchema?>();

   /// <summary>
   /// Tells two schemas apart by what they are, not by what they say. Two schemas describing the
   /// same thing are still two schemas, and the walk has to visit both.
   /// </summary>
   private sealed class SchemaIdentity : IEqualityComparer<OpenApiSchema>
   {
      public static IEqualityComparer<OpenApiSchema> Comparer { get; } = new SchemaIdentity();

      public bool Equals(OpenApiSchema? left, OpenApiSchema? right) => ReferenceEquals(left, right);

      public int GetHashCode(OpenApiSchema schema)
         => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(schema);
   }

   private static void Collect(OpenApiSchema? schema, IDictionary<string, OpenApiSchema>? declarations,
      IDictionary<string, OpenApiSchema> found, ISet<OpenApiSchema> visited)
   {
      if (schema is null || !visited.Add(schema))
      {
         return;
      }

      if (schema.GetDeclaredSchema() is { Reference.Id: { Length: > 0 } name } named)
      {
         // A reference is not always resolved into the schema it points at, a $ref under a response
         // header keeps the name of the schema and nothing else, while the components section of
         // the document always declares the schema itself.
         var declared = Declared(declarations, name) ?? named;
         if (found.TryAdd(name, declared) && !ReferenceEquals(declared, schema))
         {
            Collect(declared, declarations, found, visited);
         }
      }

      // A wrapper around a reference describes nothing itself, everything is on the schema it
      // wraps, and that schema is walked here instead of a second time on its own.
      var effective = schema.GetEffectiveSchema();
      if (!ReferenceEquals(effective, schema))
      {
         visited.Add(effective);
      }

      Collect(effective.Items, declarations, found, visited);
      Collect(effective.AdditionalProperties, declarations, found, visited);

      foreach (var composed in effective.AllOf.Concat(effective.AnyOf).Concat(effective.OneOf))
      {
         Collect(composed, declarations, found, visited);
      }

      foreach (var property in effective.Properties.Values)
      {
         Collect(property, declarations, found, visited);
      }
   }

   private static OpenApiSchema? Declared(IDictionary<string, OpenApiSchema>? declarations, string name)
      => declarations is not null && declarations.TryGetValue(name, out var declared) ? declared : null;
}
