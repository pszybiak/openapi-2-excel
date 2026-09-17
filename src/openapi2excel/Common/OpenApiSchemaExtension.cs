using Microsoft.OpenApi.Any;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Lang;
using System.Text;

namespace openapi2excel.core.Common;

internal static class OpenApiSchemaExtension
{
   /// <summary>
   /// The schema that carries the type of a property. A generator that needs to add a keyword to a
   /// referenced schema, typically "nullable: true", cannot put it next to the $ref, so it wraps
   /// the reference in a single element allOf. That wrapper has no type of its own and hides the
   /// name of the referenced schema.
   /// </summary>
   public static OpenApiSchema GetEffectiveSchema(this OpenApiSchema schema)
   {
      // A recursive schema wrapped in itself would loop, so the unwrapping is bounded.
      const int maxWrapperDepth = 20;

      var effective = schema;
      for (var depth = 0; depth < maxWrapperDepth && effective.Reference is null; depth++)
      {
         var wrapped = GetWrappedSchema(effective);
         if (wrapped is null)
         {
            break;
         }

         effective = wrapped;
      }

      return effective;
   }

   private static OpenApiSchema? GetWrappedSchema(OpenApiSchema schema)
   {
      // Anything of its own makes it a schema, not a wrapper.
      if (schema.Type is not null || schema.Items is not null || schema.Properties.Any())
      {
         return null;
      }

      var wrapped = new[] { schema.AllOf, schema.AnyOf, schema.OneOf }
         .Where(composition => composition.Count == 1)
         .Select(composition => composition[0])
         .ToList();

      // Two of them at once is a composition, and no single schema describes it.
      return wrapped.Count == 1 ? wrapped[0] : null;
   }

   public static string GetTypeDescription(this OpenApiSchema schema, Translation translation)
   {
      var effective = schema.GetEffectiveSchema();
      return effective.Type switch
      {
         "object" => translation.ObjectType,
         "array" => translation.ArrayType,
         null => translation.ObjectType,
         _ => effective.Type
      };
   }

   /// <summary>
   /// The schema under the name this document declares it with, or null when it declares none: a
   /// schema written in place, or a reference to another document, which this one never read and
   /// cannot describe.
   /// </summary>
   public static OpenApiSchema? GetDeclaredSchema(this OpenApiSchema schema)
   {
      var effective = schema.GetEffectiveSchema();
      return effective.Reference is { IsExternal: false, Id.Length: > 0 } ? effective : null;
   }

   /// <summary>
   /// The schema named by the Object type column, or null when that column names nothing this
   /// document declares. An array is named by what it holds, the way
   /// <see cref="GetObjectDescription"/> describes it.
   /// </summary>
   public static OpenApiSchema? GetNamedSchema(this OpenApiSchema schema)
   {
      // An array of arrays is legal, and one wrapped in itself would loop.
      const int maxArrayDepth = 20;

      var current = schema;
      for (var depth = 0; depth < maxArrayDepth; depth++)
      {
         var effective = current.GetEffectiveSchema();
         if (!"array".Equals(effective.Type, StringComparison.Ordinal))
         {
            return effective.GetDeclaredSchema();
         }

         if (effective.Items is null)
         {
            return null;
         }

         current = effective.Items;
      }

      return null;
   }

   /// <summary>
   /// Name of the schema the Object type column names, the one the objects worksheet documents in a
   /// section of its own. Null when the column names nothing that worksheet describes.
   /// </summary>
   public static string? GetSchemaName(this OpenApiSchema schema)
      => schema.GetNamedSchema()?.Reference?.Id;

   /// <summary>
   /// Name of the schema itself, as opposed to the name of what it holds: a list of pets is named
   /// by the name the document gives the list, and by nothing when the list is written in place.
   /// </summary>
   public static string? GetOwnSchemaName(this OpenApiSchema schema)
      => schema.GetDeclaredSchema()?.Reference?.Id;

   public static string GetObjectDescription(this OpenApiSchema schema, Translation translation)
   {
      var effective = schema.GetEffectiveSchema();
      return effective.Type switch
      {
         "object" => effective.Reference is null ? translation.ObjectType : effective.Reference.Id,
         // An array is named by what it holds, and an array holding nothing the document describes
         // names nothing. The Type column already says it is an array.
         "array" => effective.Items is null ? string.Empty : effective.Items.GetObjectDescription(translation),
         null => effective.Reference is null ? translation.ObjectType : effective.Reference.Id,
         _ => effective.Reference is null ? "" : effective.Reference.Id
      };
   }

   public static string GetPropertyDescription(this OpenApiSchema schema)
   {
      if (string.IsNullOrEmpty(schema.Description))
         return schema.Description;

      return schema.Description.StartsWith('\'') ? "'" + schema.Description : schema.Description;
   }

   public static string GetPropertyLengthDescription(this OpenApiSchema schema, Translation translation)
   {
      var culture = translation.GetCulture();
      StringBuilder propertyTypeDescription = new();
      if (schema.MinLength is not null)
      {
         propertyTypeDescription.Append(schema.MinLength.Value.ToString(culture));
      }

      if (schema.MinLength is not null && !(schema.MinLength is null && schema.MaxLength is not null))
      {
         propertyTypeDescription.Append("..");
      }

      if (schema.MaxLength is not null)
      {
         propertyTypeDescription.Append(schema.MaxLength.Value.ToString(culture));
      }

      return propertyTypeDescription.ToString();
   }

   public static string GetPropertyRangeDescription(this OpenApiSchema schema, Translation translation)
   {
      var culture = translation.GetCulture();
      StringBuilder propertyTypeDescription = new();
      if (schema.Minimum is not null)
      {
         var sign = schema.ExclusiveMinimum is null or false ? "[" : "(";
         propertyTypeDescription.Append(sign + schema.Minimum.Value.ToString(culture));
      }
      else if (schema.Maximum is not null)
      {
         propertyTypeDescription.Append("(..");
      }

      if (schema.Minimum is not null || schema.Maximum is not null)
      {
         propertyTypeDescription.Append(';');
      }

      if (schema.Maximum is not null)
      {
         var sign = schema.ExclusiveMaximum is null or false ? "]" : ")";
         propertyTypeDescription.Append(schema.Maximum.Value.ToString(culture) + sign);
      }
      else if (schema.Minimum is not null)
      {
         propertyTypeDescription.Append("..)");
      }

      return propertyTypeDescription.ToString();
   }

   public static string GetEnumDescription(this OpenApiSchema schema, Translation translation)
   {
      var culture = translation.GetCulture();
      return !schema.Enum.Any() ? string.Empty : string.Join(", ", schema.Enum.Select(GetEnumValue));

      string GetEnumValue(IOpenApiAny value)
      {
         if (value is not IOpenApiPrimitive)
            return "";

         return value switch
         {
            OpenApiString val => val.Value,
            OpenApiInteger val => val.Value.ToString(),
            OpenApiBoolean val => val.Value.ToString(),
            OpenApiByte val => val.Value.ToString(),
            OpenApiDate val => val.Value.ToString("d", culture),
            OpenApiDateTime val => val.Value.ToString(culture),
            OpenApiDouble val => val.Value.ToString(culture),
            OpenApiFloat val => val.Value.ToString(culture),
            OpenApiLong val => val.Value.ToString(culture),
            OpenApiPassword val => val.Value,
            _ => ""
         };
      }
   }

   public static string GetExampleDescription(this OpenApiSchema schema, Translation translation)
   {
      var culture = translation.GetCulture();
      if (schema.Example == null)
      {
         return string.Empty;
      }

      if (schema.Example is not IOpenApiPrimitive)
      {
         // TODO: add complex example
         return "";
      }
      return schema.Example switch
      {
         OpenApiString val => val.Value,
         OpenApiInteger val => val.Value.ToString(),
         OpenApiBoolean val => val.Value.ToString(),
         OpenApiByte val => val.Value.ToString(),
         OpenApiDate val => val.Value.ToString("d", culture),
         OpenApiDateTime val => val.Value.ToString(culture),
         OpenApiDouble val => val.Value.ToString(culture),
         OpenApiFloat val => val.Value.ToString(culture),
         OpenApiLong val => val.Value.ToString(culture),
         OpenApiPassword val => val.Value,
         _ => ""
      };
   }

   public static string GetDefaultDescription(this OpenApiSchema schema, Translation translation)
   {
      var culture = translation.GetCulture();
      if (schema.Default == null)
      {
         return string.Empty;
      }

      if (schema.Default is not IOpenApiPrimitive)
      {
         // TODO: add complex default value
         return "";
      }
      return schema.Default switch
      {
         OpenApiString val => val.Value,
         OpenApiInteger val => val.Value.ToString(),
         OpenApiBoolean val => val.Value.ToString(),
         OpenApiByte val => val.Value.ToString(),
         OpenApiDate val => val.Value.ToString("d", culture),
         OpenApiDateTime val => val.Value.ToString(culture),
         OpenApiDouble val => val.Value.ToString(culture),
         OpenApiFloat val => val.Value.ToString(culture),
         OpenApiLong val => val.Value.ToString(culture),
         OpenApiPassword val => val.Value,
         _ => ""
      };
   }
}