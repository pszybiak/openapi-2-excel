using Microsoft.OpenApi.Any;
using Microsoft.OpenApi.Models;
using System.Globalization;
using System.Text;

namespace openapi2excel.core.Common;

internal static class OpenApiSchemaExtension
{
   public static string GetTypeDescription(this OpenApiSchema schema)
   {
      return schema.Type switch
      {
         "object" => "object",
         "array" => "array",
         null => "object",
         _ => schema.Type
      };
   }

   public static string GetObjectDescription(this OpenApiSchema schema)
   {
      return schema.Type switch
      {
         "object" => schema.Reference is null ? "object" : schema.Reference.Id,
         "array" => schema.Items.GetObjectDescription(),
         null => schema.AllOf.Any() ? "All of (" + string.Join(", ", schema.AllOf.Select(GetObjectDescription)) + ")" : "object",
         _ => ""
      };
   }

   public static string GetPropertyDescription(this OpenApiSchema schema)
   {
      if (string.IsNullOrEmpty(schema.Description))
         return schema.Description;

      return schema.Description.StartsWith('\'') ? "'" + schema.Description : schema.Description;
   }

   public static string GetPropertyLengthDescription(this OpenApiSchema schema)
   {
      StringBuilder propertyTypeDescription = new();
      if (schema.MinLength is not null)
      {
         propertyTypeDescription.Append($"{schema.MinLength}");
      }

      if (schema.MinLength is not null && !(schema.MinLength is null && schema.MaxLength is not null))
      {
         propertyTypeDescription.Append("..");
      }

      if (schema.MaxLength is not null)
      {
         propertyTypeDescription.Append(schema.MaxLength);
      }

      return propertyTypeDescription.ToString();
   }

   public static string GetPropertyRangeDescription(this OpenApiSchema schema)
   {
      StringBuilder propertyTypeDescription = new();
      if (schema.Minimum is not null)
      {
         var sign = schema.ExclusiveMinimum is null or false ? "[" : "(";
         propertyTypeDescription.Append($"{sign}{schema.Minimum}");
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
         propertyTypeDescription.Append($"{schema.Maximum}{sign}");
      }
      else if (schema.Minimum is not null)
      {
         propertyTypeDescription.Append("..)");
      }

      return propertyTypeDescription.ToString();
   }

   public static string GetEnumDescription(this OpenApiSchema schema)
   {
      return !schema.Enum.Any() ? string.Empty : string.Join(", ", schema.Enum.Select(GetEnumValue));

      static string GetEnumValue(IOpenApiAny value)
      {
         if (value is not IOpenApiPrimitive)
            return "";

         return value switch
         {
            OpenApiString val => val.Value,
            OpenApiInteger val => val.Value.ToString(),
            OpenApiBoolean val => val.Value.ToString(),
            OpenApiByte val => val.Value.ToString(),
            OpenApiDate val => val.Value.ToShortDateString(),
            OpenApiDateTime val => val.Value.ToString(CultureInfo.CurrentCulture),
            OpenApiDouble val => val.Value.ToString(CultureInfo.CurrentCulture),
            OpenApiFloat val => val.Value.ToString(CultureInfo.CurrentCulture),
            OpenApiLong val => val.Value.ToString(CultureInfo.CurrentCulture),
            OpenApiPassword val => val.Value,
            _ => ""
         };
      }
   }

   public static string GetExampleDescription(this OpenApiSchema schema)
   {
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
         OpenApiDate val => val.Value.ToShortDateString(),
         OpenApiDateTime val => val.Value.ToString(CultureInfo.CurrentCulture),
         OpenApiDouble val => val.Value.ToString(CultureInfo.CurrentCulture),
         OpenApiFloat val => val.Value.ToString(CultureInfo.CurrentCulture),
         OpenApiLong val => val.Value.ToString(CultureInfo.CurrentCulture),
         OpenApiPassword val => val.Value,
         _ => ""
      };
   }

   public static string GetDefaultDescription(this OpenApiSchema schema)
   {
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
         OpenApiDate val => val.Value.ToShortDateString(),
         OpenApiDateTime val => val.Value.ToString(CultureInfo.CurrentCulture),
         OpenApiDouble val => val.Value.ToString(CultureInfo.CurrentCulture),
         OpenApiFloat val => val.Value.ToString(CultureInfo.CurrentCulture),
         OpenApiLong val => val.Value.ToString(CultureInfo.CurrentCulture),
         OpenApiPassword val => val.Value,
         _ => ""
      };
   }
}