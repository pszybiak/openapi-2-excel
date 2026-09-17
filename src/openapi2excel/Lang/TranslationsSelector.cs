using System.Reflection;
using System.Text.Json;

namespace openapi2excel.core.Lang;

/// <summary>
/// Loads a translation, either one of the languages shipped with the library or a json file of the
/// same shape from disk.
/// </summary>
public class TranslationsSelector
{
   public const string DefaultLanguage = "en";

   private const string ResourcePrefix = "openapi2excel.core.Lang.";
   private const string ResourceSuffix = ".json";

   // Reading and deserializing the embedded resource on every label lookup costs more than
   // generating the whole document, so the default translation is built once.
   private static readonly Lazy<Translation> DefaultTranslation =
      new(() => Deserialize(ReadEmbedded(DefaultLanguage), DefaultLanguage));

   private static readonly Lazy<IReadOnlyList<string>> AvailableLanguagesLoader = new(FindAvailableLanguages);

   private static readonly Dictionary<string, string> EmbeddedContents = new(StringComparer.OrdinalIgnoreCase);

   /// <summary>The language used when none is chosen.</summary>
   public static Translation Default => DefaultTranslation.Value;

   /// <summary>Languages shipped with the library, as language codes.</summary>
   public static IReadOnlyList<string> AvailableLanguages => AvailableLanguagesLoader.Value;

   /// <summary>
   /// Loads the translation named by a language code, "pl", or by a path to a json file of the same
   /// shape. Labels the translation does not define are taken from <see cref="Default"/> and listed
   /// in the result, so an incomplete translation still produces a complete document.
   /// </summary>
   /// <exception cref="InvalidLanguageException">
   /// The language is neither a shipped language code nor a readable json file.
   /// </exception>
   public static TranslationResult Load(string? languageOrPath)
   {
      if (string.IsNullOrWhiteSpace(languageOrPath))
      {
         return new TranslationResult(Default, []);
      }

      var (content, source) = LooksLikeFile(languageOrPath!)
         ? (ReadFile(languageOrPath!), languageOrPath!)
         : (ReadEmbedded(languageOrPath!), languageOrPath!);

      var translation = Deserialize(content, source);
      return new TranslationResult(translation, FillMissingLabelsFromDefault(translation));
   }

   /// <summary>
   /// Loads the translation named by a language code or by a path to a json file, without telling
   /// which labels it had to take from the default language. Use <see cref="Load"/> to learn that.
   /// </summary>
   public Translation GetTranslation(string language) => Load(language).Translation;

   private static bool LooksLikeFile(string languageOrPath)
      => languageOrPath.EndsWith(ResourceSuffix, StringComparison.OrdinalIgnoreCase)
         || languageOrPath.IndexOfAny([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar]) >= 0;

   private static string ReadFile(string path)
      => File.Exists(path)
         ? File.ReadAllText(path)
         : throw new InvalidLanguageException($"Translation file not found: {path}.");

   private static string ReadEmbedded(string language)
   {
      lock (EmbeddedContents)
      {
         if (EmbeddedContents.TryGetValue(language, out var cached))
         {
            return cached;
         }
      }

      var name = AvailableLanguages.FirstOrDefault(
         available => available.Equals(language, StringComparison.OrdinalIgnoreCase))
         ?? throw new InvalidLanguageException(
            $"Unknown language: {language}. Available languages: {string.Join(", ", AvailableLanguages)}. " +
            "A path to a json file with a translation is accepted as well.");

      using var stream = typeof(TranslationsSelector).Assembly.GetManifestResourceStream(ResourcePrefix + name + ResourceSuffix)
         ?? throw new InvalidLanguageException($"Translation resource not found for language {name}.");
      using var reader = new StreamReader(stream);
      var content = reader.ReadToEnd();

      lock (EmbeddedContents)
      {
         EmbeddedContents[name] = content;
      }

      return content;
   }

   private static Translation Deserialize(string json, string source)
   {
      try
      {
         return JsonSerializer.Deserialize<Translation>(json)
            ?? throw new InvalidLanguageException($"Translation is empty: {source}.");
      }
      catch (JsonException exception)
      {
         throw new InvalidLanguageException($"Translation is not a valid json file: {source}.", exception);
      }
   }

   /// <summary>
   /// Copies every label the translation leaves empty from the default language and returns their
   /// names. A translation written for an older version of the tool keeps working, and a label it
   /// does not know shows up in English instead of leaving an empty cell in the document.
   /// </summary>
   internal static IReadOnlyList<string> FillMissingLabelsFromDefault(Translation translation)
   {
      var missing = new List<string>();
      foreach (var label in Labels)
      {
         if (!string.IsNullOrEmpty((string?)label.GetValue(translation)))
         {
            continue;
         }

         label.SetValue(translation, label.GetValue(Default));
         missing.Add(label.Name);
      }

      return missing;
   }

   private static IReadOnlyList<string> FindAvailableLanguages()
      => typeof(TranslationsSelector).Assembly.GetManifestResourceNames()
         .Where(name => name.StartsWith(ResourcePrefix, StringComparison.Ordinal)
                        && name.EndsWith(ResourceSuffix, StringComparison.Ordinal))
         .Select(name => name.Substring(ResourcePrefix.Length,
            name.Length - ResourcePrefix.Length - ResourceSuffix.Length))
         .OrderBy(name => !name.Equals(DefaultLanguage, StringComparison.OrdinalIgnoreCase))
         .ThenBy(name => name, StringComparer.OrdinalIgnoreCase)
         .ToList();

   internal static IReadOnlyList<PropertyInfo> Labels { get; } = typeof(Translation)
      .GetProperties(BindingFlags.Public | BindingFlags.Instance)
      .Where(property => property.PropertyType == typeof(string) && property.CanRead && property.CanWrite)
      .OrderBy(property => property.Name, StringComparer.Ordinal)
      .ToList();
}
