namespace openapi2excel.core.Sanitization;

/// <summary>
/// A specification violation the sanitizer knows how to repair.
/// </summary>
public enum SanitizationRule
{
   /// <summary>
   /// Two paths differ only in the name of their parameters, which the specification forbids.
   /// The operations of the second path are moved to the first one.
   /// </summary>
   DuplicatePathSignature,

   /// <summary>
   /// A path parameter does not match the path template and exactly one placeholder of that
   /// template is undeclared, so the parameter is renamed to it.
   /// </summary>
   PathParameterRenamedToTemplate,

   /// <summary>
   /// A path parameter does not match the path template and every placeholder of that template
   /// is already declared, so the parameter is turned into a query parameter.
   /// </summary>
   PathParameterMovedToQuery
}

/// <summary>
/// A single repair applied to the document.
/// </summary>
public sealed class SanitizationCorrection(SanitizationRule rule, string location, string description)
{
   public SanitizationRule Rule { get; } = rule;

   /// <summary>
   /// Path, or path and operation, the correction was applied to.
   /// </summary>
   public string Location { get; } = location;

   public string Description { get; } = description;

   public override string ToString() => $"{Location}: {Description}";
}

/// <summary>
/// A violation the sanitizer detected but refused to repair, because repairing it would lose
/// information. The document is still rejected because of it.
/// </summary>
public sealed class SanitizationProblem(SanitizationRule rule, string location, string description)
{
   public SanitizationRule Rule { get; } = rule;

   public string Location { get; } = location;

   public string Description { get; } = description;

   public override string ToString() => $"{Location}: {Description}";
}

/// <summary>
/// What the sanitizer changed in the input document.
/// </summary>
public sealed class SanitizationReport(
   IReadOnlyList<SanitizationCorrection> corrections,
   IReadOnlyList<SanitizationProblem> skippedProblems)
{
   public IReadOnlyList<SanitizationCorrection> Corrections { get; } = corrections;

   public IReadOnlyList<SanitizationProblem> SkippedProblems { get; } = skippedProblems;

   public bool HasChanges => Corrections.Count > 0 || SkippedProblems.Count > 0;
}
