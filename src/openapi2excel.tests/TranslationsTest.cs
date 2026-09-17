using System.Globalization;
using System.Reflection;
using System.Text.Json;
using openapi2excel.core.Lang;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// Gate for every language shipped with the library. A label added to <see cref="Translation"/>
   /// and forgotten in a language file, or a key left in a file after the label was renamed, fails
   /// here instead of leaving a blank cell in a document nobody reads carefully.
   /// </summary>
   public class TranslationsTest
   {
      private const string ResourcePrefix = "openapi2excel.core.Lang.";

      private static readonly IReadOnlyList<string> Labels = typeof(Translation)
         .GetProperties(BindingFlags.Public | BindingFlags.Instance)
         .Where(property => property.PropertyType == typeof(string))
         .Select(property => property.Name)
         .OrderBy(name => name, StringComparer.Ordinal)
         .ToList();

      public static TheoryData<string> ShippedLanguages()
      {
         var languages = new TheoryData<string>();
         foreach (var language in TranslationsSelector.AvailableLanguages)
         {
            languages.Add(language);
         }

         return languages;
      }

      [Fact]
      public void English_and_polish_are_shipped_and_english_comes_first()
      {
         Assert.Equal("en", TranslationsSelector.DefaultLanguage);
         Assert.Equal("en", TranslationsSelector.AvailableLanguages[0]);
         Assert.Contains("pl", TranslationsSelector.AvailableLanguages);
      }

      [Theory]
      [MemberData(nameof(ShippedLanguages))]
      public void Shipped_language_defines_every_label_and_nothing_else(string language)
      {
         var keys = ReadKeys(language);

         Assert.Equal(Labels, keys.OrderBy(key => key, StringComparer.Ordinal));
      }

      [Theory]
      [MemberData(nameof(ShippedLanguages))]
      public void Shipped_language_needs_nothing_from_the_default_language(string language)
      {
         var result = TranslationsSelector.Load(language);

         Assert.Empty(result.MissingLabels);
         Assert.True(result.IsComplete);
      }

      [Theory]
      [MemberData(nameof(ShippedLanguages))]
      public void Shipped_language_names_a_culture_the_machine_knows(string language)
      {
         var translation = TranslationsSelector.Load(language).Translation;

         Assert.False(string.IsNullOrWhiteSpace(translation.Culture));
         var culture = CultureInfo.GetCultureInfo(translation.Culture);
         Assert.Equal(culture, translation.GetType()
            .GetMethod("GetCulture", BindingFlags.NonPublic | BindingFlags.Instance)!
            .Invoke(translation, null));
      }

      [Theory]
      [MemberData(nameof(ShippedLanguages))]
      public void Shipped_language_keeps_the_placeholders_of_the_composed_labels(string language)
      {
         var translation = TranslationsSelector.Load(language).Translation;

         // These labels are format strings, a translation that drops the placeholder silently drops
         // the media type or the http code from the document.
         Assert.Contains("{0}", translation.BodyFormat);
         Assert.Contains("{0}", translation.ResponseHttpCode);
         Assert.Contains("{0}", translation.ResponseWithDescription);
         Assert.Contains("{1}", translation.ResponseWithDescription);
         Assert.Contains("{0}", translation.BodyFormatWithObject);
         Assert.Contains("{1}", translation.BodyFormatWithObject);
         Assert.Contains("{0}", translation.InfoOperationWithId);
         Assert.Contains("{1}", translation.InfoOperationWithId);
         Assert.Contains("{0}", translation.AllOfType);
         Assert.Contains("{0}", translation.OneOfType);
         Assert.Contains("{0}", translation.AnyOfType);
      }

      [Theory]
      [MemberData(nameof(ShippedLanguages))]
      public void Shipped_language_translates_every_label(string language)
      {
         if (language == TranslationsSelector.DefaultLanguage)
         {
            return;
         }

         var translation = TranslationsSelector.Load(language).Translation;
         var english = TranslationsSelector.Default;

         // Labels that are the same word in both languages, or not words at all.
         string[] sameInEveryLanguage =
         [
            nameof(Translation.FieldFormat), nameof(Translation.ParameterSerialization),
            nameof(Translation.InfoSheetName), nameof(Translation.ResponseWithDescription),
            nameof(Translation.InfoTag), nameof(Translation.BodyFormatWithObject),
            nameof(Translation.InfoOperationWithId)
         ];

         var untranslated = typeof(Translation)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(property => property.PropertyType == typeof(string))
            .Where(property => !sameInEveryLanguage.Contains(property.Name))
            .Where(property => property.Name != nameof(Translation.Culture))
            .Where(property => (string?)property.GetValue(translation) == (string?)property.GetValue(english))
            .Select(property => property.Name)
            .ToList();

         Assert.Empty(untranslated);
      }

      private static IReadOnlyCollection<string> ReadKeys(string language)
      {
         using var stream = typeof(Translation).Assembly
            .GetManifestResourceStream($"{ResourcePrefix}{language}.json")!;
         using var reader = new StreamReader(stream);

         return JsonSerializer.Deserialize<Dictionary<string, string>>(reader.ReadToEnd())!.Keys;
      }
   }
}
