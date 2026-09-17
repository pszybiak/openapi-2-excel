using ClosedXML.Excel;

namespace openapi2excel.core.Builders;

internal static class WorksheetName
{
   /// <summary>Characters Excel refuses in the name of a worksheet.</summary>
   private static readonly char[] Forbidden = [':', '\\', '/', '?', '*', '[', ']'];

   /// <summary>
   /// A name Excel accepts and no worksheet of the workbook uses yet, at most
   /// <paramref name="maxLength"/> characters long. Excel allows 31 characters in a worksheet name,
   /// refuses a handful of characters in one, and refuses two worksheets named the same, whatever
   /// the case. A specification is free to repeat an operationId and to put any character in one,
   /// and so is a translation written by hand. A name already taken gets a number, and is trimmed
   /// so that the number still fits in the limit.
   /// </summary>
   public static string Unique(IXLWorkbook workbook, string name, int maxLength)
   {
      name = Accepted(name);
      if (name.Length > maxLength)
      {
         name = name[..maxLength];
      }

      var number = 2;
      var candidate = name;
      while (workbook.Worksheets.Any(worksheet
                => worksheet.Name.Equals(candidate, StringComparison.CurrentCultureIgnoreCase)))
      {
         var suffix = $"_{number++}";
         candidate = name[..Math.Min(name.Length, maxLength - suffix.Length)] + suffix;
      }

      return candidate;
   }

   /// <summary>
   /// The name without the characters Excel refuses, and never empty: a name made of nothing but
   /// those characters would leave the worksheet nameless.
   /// </summary>
   private static string Accepted(string name)
   {
      var accepted = new string(name
         .Select(character => Forbidden.Contains(character) ? '-' : character)
         .ToArray())
         .Trim(' ', '\'');

      return accepted.Length > 0 ? accepted : "-";
   }
}
