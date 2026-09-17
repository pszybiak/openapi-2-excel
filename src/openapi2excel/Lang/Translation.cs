using System.Globalization;

namespace openapi2excel.core.Lang;

public class Translation
{
   private CultureInfo? _culture;

   /// <summary>
   /// Culture used to format the values taken from the specification: numbers, dates, ranges. It
   /// belongs to the translation so that the same specification always produces the same document,
   /// whatever the culture of the machine that generates it.
   /// </summary>
   public string Culture { get; set; } = null!;

   public string Yes { get; set; } = null!;
   public string No { get; set; } = null!;
   public string InfoSheetName { get; set; } = null!;
   public string InfoTitle { get; set; } = null!;
   public string InfoVersion { get; set; } = null!;
   public string InfoDescription { get; set; } = null!;

   /// <summary>Header of the index of operations on the info worksheet.</summary>
   public string InfoOperations { get; set; } = null!;

   /// <summary>Column of the index of operations holding the tag the operation is grouped under.</summary>
   public string InfoTag { get; set; } = null!;

   /// <summary>Group of the index of operations holding the operations carrying no tag.</summary>
   public string InfoUntagged { get; set; } = null!;

   /// <summary>
   /// Takes the type of an operation and the id naming it, for example "{0}: {1}". The index of the
   /// operations states both in one column.
   /// </summary>
   public string InfoOperationWithId { get; set; } = null!;

   /// <summary>Name of the last worksheet, the one documenting the objects of the specification.</summary>
   public string ObjectsSheetName { get; set; } = null!;

   /// <summary>Header of the worksheet documenting the objects.</summary>
   public string ObjectsHeader { get; set; } = null!;
   public string FieldName { get; set; } = null!;
   public string FieldDescription { get; set; } = null!;
   public string FieldObjectType { get; set; } = null!;
   public string FieldType { get; set; } = null!;
   public string FieldFormat { get; set; } = null!;
   public string FieldLength { get; set; } = null!;
   public string FieldRequired { get; set; } = null!;
   public string FieldNullable { get; set; } = null!;
   public string FieldRange { get; set; } = null!;
   public string FieldPattern { get; set; } = null!;
   public string FieldEnum { get; set; } = null!;
   public string FieldDeprecated { get; set; } = null!;
   public string FieldDefault { get; set; } = null!;
   public string FieldExample { get; set; } = null!;
   public string ArrayType { get; set; } = null!;
   public string ObjectType { get; set; } = null!;

   /// <summary>
   /// Takes the names of the schemas a schema is written as, for example "All of ({0})". The
   /// Object type column describes a schema written as all of several that way.
   /// </summary>
   public string AllOfType { get; set; } = null!;

   /// <summary>
   /// Takes the names of the schemas a schema is written as, for example "One of ({0})". The
   /// Object type column describes a schema written as one of several that way.
   /// </summary>
   public string OneOfType { get; set; } = null!;

   /// <summary>
   /// Takes the names of the schemas a schema is written as, for example "Any of ({0})". The
   /// Object type column describes a schema written as any of several that way.
   /// </summary>
   public string AnyOfType { get; set; } = null!;
   public string OperationHeader { get; set; } = null!;
   public string OperationType { get; set; } = null!;
   public string OperationId { get; set; } = null!;
   public string OperationPath { get; set; } = null!;
   public string OperationPathDescription { get; set; } = null!;
   public string OperationPathSummary { get; set; } = null!;
   public string OperationDescription { get; set; } = null!;
   public string OperationSummary { get; set; } = null!;
   public string OperationDeprecated { get; set; } = null!;

   /// <summary>Takes the media type, for example "Body format: {0}".</summary>
   public string BodyFormat { get; set; } = null!;

   /// <summary>
   /// Takes the body format label and the name of the object the body is, for example "{0} ({1})".
   /// </summary>
   public string BodyFormatWithObject { get; set; } = null!;

   public string ArrayPlaceholder { get; set; } = null!;
   public string ValuePlaceholder { get; set; } = null!;
   public string RequestHeader { get; set; } = null!;
   public string ResponseHeader { get; set; } = null!;
   public string ParametersHeader { get; set; } = null!;
   public string ParameterName { get; set; } = null!;
   public string ParameterLocation { get; set; } = null!;
   public string ParameterSerialization { get; set; } = null!;
   public string ResponseHeaders { get; set; } = null!;
   public string ResponseDefault { get; set; } = null!;

   /// <summary>Takes the http code, for example "Response HttpCode: {0}".</summary>
   public string ResponseHttpCode { get; set; } = null!;

   /// <summary>Takes the response label and its description, for example "{0}: {1}".</summary>
   public string ResponseWithDescription { get; set; } = null!;

   /// <summary>
   /// <see cref="Culture"/> as a <see cref="CultureInfo"/>. A name no machine knows falls back to
   /// the invariant culture instead of failing the whole document.
   /// </summary>
   internal CultureInfo GetCulture()
   {
      if (_culture is not null)
      {
         return _culture;
      }

      try
      {
         _culture = string.IsNullOrWhiteSpace(Culture)
            ? CultureInfo.InvariantCulture
            : CultureInfo.GetCultureInfo(Culture);
      }
      catch (CultureNotFoundException)
      {
         _culture = CultureInfo.InvariantCulture;
      }

      return _culture;
   }
}
