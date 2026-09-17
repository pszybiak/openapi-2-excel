using openapi2excel.core;
using openapi2excel.core.Lang;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// Choosing a language: a code of a shipped language, or a path to a json file for a language
   /// nobody has contributed yet.
   /// </summary>
   public class LanguageSelectionTest
   {
      [Theory]
      [InlineData(null)]
      [InlineData("")]
      [InlineData("   ")]
      public void No_language_means_the_default_one(string? language)
      {
         var result = TranslationsSelector.Load(language);

         Assert.Same(TranslationsSelector.Default, result.Translation);
         Assert.True(result.IsComplete);
      }

      [Theory]
      [InlineData("pl")]
      [InlineData("PL")]
      [InlineData("Pl")]
      public void Language_code_is_case_insensitive(string language)
         => Assert.Equal("Informacje", TranslationsSelector.Load(language).Translation.InfoSheetName);

      [Fact]
      public void Default_translation_is_read_once()
      {
         // Every label of every row of every worksheet goes through this, so it has to be the same
         // instance and not a fresh read of the embedded resource.
         Assert.Same(TranslationsSelector.Default, TranslationsSelector.Default);
      }

      [Fact]
      public void Loaded_translation_is_a_fresh_instance_that_cannot_poison_the_next_load()
      {
         var first = TranslationsSelector.Load("pl").Translation;
         first.InfoSheetName = "wrecked";

         Assert.Equal("Informacje", TranslationsSelector.Load("pl").Translation.InfoSheetName);
      }

      [Fact]
      public void Unknown_language_lists_what_is_available()
      {
         var exception = Assert.Throws<InvalidLanguageException>(() => TranslationsSelector.Load("kl"));

         Assert.Contains("Unknown language: kl", exception.Message);
         foreach (var language in TranslationsSelector.AvailableLanguages)
         {
            Assert.Contains(language, exception.Message);
         }
      }

      [Fact]
      public void Translation_can_come_from_a_file()
      {
         var file = WriteTranslation("""
            {
              "Culture": "cs-CZ",
              "Yes": "Ano",
              "No": "Ne",
              "InfoSheetName": "Informace",
              "FieldName": "Jmeno"
            }
            """);
         try
         {
            var result = TranslationsSelector.Load(file);

            Assert.Equal("Ano", result.Translation.Yes);
            Assert.Equal("Informace", result.Translation.InfoSheetName);
            Assert.Equal("Jmeno", result.Translation.FieldName);
         }
         finally
         {
            File.Delete(file);
         }
      }

      [Fact]
      public void Labels_a_file_does_not_define_come_from_the_default_language_and_are_reported()
      {
         var file = WriteTranslation("""
            { "Yes": "Ano", "No": "Ne" }
            """);
         try
         {
            var result = TranslationsSelector.Load(file);

            Assert.False(result.IsComplete);
            Assert.Equal("Ano", result.Translation.Yes);
            Assert.Equal(TranslationsSelector.Default.FieldName, result.Translation.FieldName);
            Assert.Equal(TranslationsSelector.Default.Culture, result.Translation.Culture);

            // Everything but the two labels the file defines.
            Assert.DoesNotContain(nameof(Translation.Yes), result.MissingLabels);
            Assert.DoesNotContain(nameof(Translation.No), result.MissingLabels);
            Assert.Contains(nameof(Translation.FieldName), result.MissingLabels);
            Assert.Contains(nameof(Translation.Culture), result.MissingLabels);
            Assert.Contains(nameof(Translation.ResponseWithDescription), result.MissingLabels);
         }
         finally
         {
            File.Delete(file);
         }
      }

      [Fact]
      public void Missing_file_is_named_in_the_error()
      {
         var exception = Assert.Throws<InvalidLanguageException>(
            () => TranslationsSelector.Load(Path.Combine(Path.GetTempPath(), "no-such-translation.json")));

         Assert.Contains("no-such-translation.json", exception.Message);
      }

      [Fact]
      public void Broken_file_is_named_in_the_error()
      {
         var file = WriteTranslation("{ this is not json");
         try
         {
            var exception = Assert.Throws<InvalidLanguageException>(() => TranslationsSelector.Load(file));

            Assert.Contains(file, exception.Message);
            Assert.IsType<System.Text.Json.JsonException>(exception.InnerException);
         }
         finally
         {
            File.Delete(file);
         }
      }

      [Fact]
      public void Obsolete_instance_api_still_works()
      {
         Assert.Equal("Informacje", new TranslationsSelector().GetTranslation("pl").InfoSheetName);
      }

      [Fact]
      public void Translation_built_by_hand_is_completed_with_the_default_language()
      {
         // Nothing forces a library consumer to go through TranslationsSelector, and a label left
         // null used to blank a cell or, once labels became format strings, throw.
         var options = new OpenApiDocumentationOptions
         {
            Translation = new Translation { InfoSheetName = "Sheet", Yes = "Y", No = "N" }
         };

         Assert.Equal("Sheet", options.Translation.InfoSheetName);
         Assert.Equal("Y", options.Translation.Yes);
         Assert.Equal(TranslationsSelector.Default.FieldName, options.Translation.FieldName);
         Assert.Equal(TranslationsSelector.Default.ResponseWithDescription, options.Translation.ResponseWithDescription);
      }

      private static string WriteTranslation(string content)
      {
         var file = Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.json");
         File.WriteAllText(file, content);
         return file;
      }
   }
}
