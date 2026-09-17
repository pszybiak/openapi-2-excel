namespace openapi2excel.core.Lang;

/// <summary>
/// A loaded translation and the labels it did not define, which were taken from the default
/// language instead.
/// </summary>
public sealed class TranslationResult(Translation translation, IReadOnlyList<string> missingLabels)
{
   public Translation Translation { get; } = translation;

   /// <summary>
   /// Names of the labels filled in from the default language, in alphabetical order. Empty when
   /// the translation is complete.
   /// </summary>
   public IReadOnlyList<string> MissingLabels { get; } = missingLabels;

   public bool IsComplete => MissingLabels.Count == 0;
}
