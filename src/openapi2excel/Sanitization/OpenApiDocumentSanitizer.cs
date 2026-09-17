using Microsoft.OpenApi.Models;
using System.Text.RegularExpressions;

namespace openapi2excel.core.Sanitization;

/// <summary>
/// Repairs the specification violations that make the reader reject an otherwise usable document.
/// Generators do produce such documents: a renamed route parameter that leaves two paths with the
/// same signature, a typo between the path template and the parameter list, a parameter left
/// behind as "in: path" after the template changed.
/// <para>
/// No operation is ever dropped. When a violation cannot be repaired without losing an operation,
/// it is reported as a skipped problem and the document stays invalid.
/// </para>
/// </summary>
internal static class OpenApiDocumentSanitizer
{
   private static readonly Regex Placeholder = new(@"\{[^{}]*\}", RegexOptions.Compiled);

   public static SanitizationReport Sanitize(OpenApiDocument document)
   {
      var corrections = new List<SanitizationCorrection>();
      var skippedProblems = new List<SanitizationProblem>();

      MergePathsWithTheSameSignature(document, corrections, skippedProblems);
      AlignPathParametersWithTemplate(document, corrections);

      return new SanitizationReport(corrections, skippedProblems);
   }

   /// <summary>
   /// "The path signature '/pet/{}' MUST be unique." Two paths that differ only in the name
   /// of a parameter are the same path, so their operations belong to a single path item.
   /// </summary>
   private static void MergePathsWithTheSameSignature(OpenApiDocument document,
      List<SanitizationCorrection> corrections, List<SanitizationProblem> skippedProblems)
   {
      var duplicatedSignatures = document.Paths.Keys
         .GroupBy(GetSignature)
         .Where(group => group.Count() > 1)
         .Select(group => group.ToList())
         .ToList();

      foreach (var paths in duplicatedSignatures)
      {
         var target = paths[0];
         var targetItem = document.Paths[target];
         var targetParameterNames = GetTemplateParameterNames(target);

         foreach (var duplicate in paths.Skip(1))
         {
            var duplicateItem = document.Paths[duplicate];
            var alreadyDefined = duplicateItem.Operations.Keys
               .Where(targetItem.Operations.ContainsKey)
               .ToList();

            if (alreadyDefined.Any())
            {
               skippedProblems.Add(new SanitizationProblem(SanitizationRule.DuplicatePathSignature, duplicate,
                  $"has the same signature as '{target}' and both define " +
                  $"{string.Join(", ", alreadyDefined.Select(operationType => operationType.ToString().ToUpperInvariant()))}, " +
                  "so the operations cannot be merged"));
               continue;
            }

            var renamedParameters = GetTemplateParameterNames(duplicate)
               .Zip(targetParameterNames, (from, to) => new { from, to })
               .Where(rename => !rename.from.Equals(rename.to, StringComparison.Ordinal))
               .ToDictionary(rename => rename.from, rename => rename.to, StringComparer.Ordinal);

            RenamePathParameters(duplicateItem.Parameters, renamedParameters);
            foreach (var operation in duplicateItem.Operations)
            {
               RenamePathParameters(operation.Value.Parameters, renamedParameters);
               targetItem.Operations[operation.Key] = operation.Value;
            }

            foreach (var parameter in duplicateItem.Parameters.Where(parameter => !IsDeclared(targetItem, parameter)))
            {
               targetItem.Parameters.Add(parameter);
            }

            document.Paths.Remove(duplicate);

            var renameDescription = renamedParameters.Any()
               ? $" (parameter {string.Join(", ", renamedParameters.Select(rename => $"'{rename.Key}' renamed to '{rename.Value}'"))})"
               : string.Empty;
            corrections.Add(new SanitizationCorrection(SanitizationRule.DuplicatePathSignature, duplicate,
               $"{string.Join(", ", duplicateItem.Operations.Keys.Select(operationType => operationType.ToString().ToUpperInvariant()))} " +
               $"moved to '{target}', which has the same signature{renameDescription}"));
         }
      }
   }

   /// <summary>
   /// "Declared path parameter \"petId\" needs to be defined as a path parameter at either the
   /// path or operation level." Either the parameter name is a typo for a placeholder nobody
   /// declared, or the parameter is not a path parameter at all.
   /// </summary>
   private static void AlignPathParametersWithTemplate(OpenApiDocument document,
      List<SanitizationCorrection> corrections)
   {
      foreach (var path in document.Paths)
      {
         var templateParameterNames = GetTemplateParameterNames(path.Key);
         if (!templateParameterNames.Any())
         {
            continue;
         }

         AlignParameters(path.Value.Parameters, templateParameterNames, [], path.Key, corrections);
         foreach (var operation in path.Value.Operations)
         {
            AlignParameters(operation.Value.Parameters, templateParameterNames, path.Value.Parameters,
               $"{operation.Key.ToString().ToUpperInvariant()} {path.Key}", corrections);
         }
      }
   }

   private static void AlignParameters(IList<OpenApiParameter> parameters, List<string> templateParameterNames,
      IList<OpenApiParameter> pathLevelParameters, string location, List<SanitizationCorrection> corrections)
   {
      var declared = pathLevelParameters.Concat(parameters)
         .Where(IsPathParameter)
         .Select(parameter => parameter.Name)
         .ToList();
      var undeclared = templateParameterNames
         .Where(name => !declared.Contains(name, StringComparer.Ordinal))
         .ToList();

      foreach (var parameter in parameters.Where(IsPathParameter).ToList())
      {
         if (templateParameterNames.Contains(parameter.Name, StringComparer.Ordinal))
         {
            continue;
         }

         if (undeclared.Count == 1)
         {
            var name = undeclared[0];
            undeclared.Clear();
            corrections.Add(new SanitizationCorrection(SanitizationRule.PathParameterRenamedToTemplate, location,
               $"path parameter '{parameter.Name}' renamed to '{name}' to match the path template"));
            parameter.Name = name;
         }
         else
         {
            corrections.Add(new SanitizationCorrection(SanitizationRule.PathParameterMovedToQuery, location,
               $"parameter '{parameter.Name}' is not in the path template, documented as a query parameter"));
            parameter.In = ParameterLocation.Query;
            parameter.Required = false;
         }
      }
   }

   private static void RenamePathParameters(IEnumerable<OpenApiParameter> parameters,
      IReadOnlyDictionary<string, string> renamedParameters)
   {
      foreach (var parameter in parameters.Where(IsPathParameter))
      {
         if (renamedParameters.TryGetValue(parameter.Name, out var name))
         {
            parameter.Name = name;
         }
      }
   }

   private static bool IsDeclared(OpenApiPathItem pathItem, OpenApiParameter parameter)
      => pathItem.Parameters.Any(declared => declared.In == parameter.In
                                             && string.Equals(declared.Name, parameter.Name, StringComparison.Ordinal));

   private static bool IsPathParameter(OpenApiParameter parameter) => parameter.In == ParameterLocation.Path;

   private static string GetSignature(string path) => Placeholder.Replace(path, "{}");

   private static List<string> GetTemplateParameterNames(string path)
      => Placeholder.Matches(path)
         .Select(match => match.Value[1..^1])
         .ToList();
}
