using openapi2excel.core.Lang;
using openapi2excel.core.Sanitization;

namespace openapi2excel.core;

public class OpenApiDocumentationOptions
{
   private Translation? _language;

   public Translation Translation
   {
      get => _language ?? TranslationsSelector.Default;
      set
      {
         // A translation built by hand, instead of loaded by TranslationsSelector, can leave labels
         // empty, and nothing downstream expects a label to be null.
         if (value is not null)
         {
            TranslationsSelector.FillMissingLabelsFromDefault(value);
         }

         _language = value;
      }
   }

   public int MaxDepth { get; set; } = 10;

   /// <summary>
   /// Repair the input document instead of rejecting it, when it violates the OpenAPI
   /// specification in a way that can be repaired without losing an operation. Enabled by default.
   /// </summary>
   public bool SanitizeDocument { get; set; } = true;

   /// <summary>
   /// Called with what the sanitizer changed, before the document is generated, and only when it
   /// changed something.
   /// </summary>
   public Action<SanitizationReport>? OnDocumentSanitized { get; set; }
}

//// TODO: Create language helper and refactor all text usage
//public class OpenApiDocumentationLanguage : Dictionary<string, string>
//{
//   public string Get(string key)
//      => TryGetValue(key, out var result) ? result : "<LANG_ERROR>";

//   public string Get(bool value)
//   {
//      if (value)
//      {
//         return TryGetValue(OpenApiDocumentationLanguageConst.Yes, out var yesResult) ? yesResult : "<LANG_ERROR>";
//      }
//      return TryGetValue(OpenApiDocumentationLanguageConst.No, out var noResult) ? noResult : "<LANG_ERROR>";
//   }

//   internal static OpenApiDocumentationLanguage Default
//      => new()
//      {
//         { OpenApiDocumentationLanguageConst.Info, "Info" },
//         { OpenApiDocumentationLanguageConst.Title, "Title" },
//         { OpenApiDocumentationLanguageConst.Version, "Version" },
//         { OpenApiDocumentationLanguageConst.Description, "Description" },

//         { OpenApiDocumentationLanguageConst.Path, "Path" },
//         { OpenApiDocumentationLanguageConst.PathSummary, "Path summary" },
//         { OpenApiDocumentationLanguageConst.PathDescription, "Path description" },

//         { OpenApiDocumentationLanguageConst.OperationType, "Operation type" },
//         { OpenApiDocumentationLanguageConst.OperationSummary, "Operation summary" },
//         { OpenApiDocumentationLanguageConst.OperationDescription, "Operation description" },
//         { OpenApiDocumentationLanguageConst.Deprecated, "Deprecated" },

//         { OpenApiDocumentationLanguageConst.Yes, "Yes" },
//         { OpenApiDocumentationLanguageConst.No, "No" }
//      };
//}

//public static class OpenApiDocumentationLanguageConst
//{
//   public const string Info = "Info";
//   public const string Title = "Title";
//   public const string Version = "Version";
//   public const string Description = "Description";

//   public const string Path = "Path";
//   public const string PathDescription = "PathDescription";
//   public const string PathSummary = "PathSummary";

//   public const string OperationType = "OperationType";
//   public const string OperationDescription = "OperationDescription";
//   public const string OperationSummary = "OperationSummary";
//   public const string Deprecated = "Deprecated";

//   public const string Yes = "Yes";
//   public const string No = "No";
//}