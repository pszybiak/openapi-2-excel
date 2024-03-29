﻿using Microsoft.OpenApi.Models;
using System.Text;

namespace OpenApi2Excel.Common;

internal static class OpenApiSchemaExtension
{
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
}