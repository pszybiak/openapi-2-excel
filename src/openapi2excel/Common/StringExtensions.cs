using openapi2excel.core.Lang;
using System.Text.RegularExpressions;

namespace openapi2excel.core.Common;

public static class StringExtensions
{
   public static string? StripHtmlTags(this string? html)
   {
      if (html is null)
      {
         return null;
      }
      html = html.Replace("<li>", "- ");
      html = Regex.Replace(html.Replace("<li>", "- "), "<.*?>", string.Empty);
      html = Regex.Replace(html, @"[\r\n]+", "\r\n");
      return html;
   }
}

public static class BoolExtensions
{
   public static string Translate(this bool value, Translation translation)
      => value ? translation.Yes : translation.No;

   public static string Translate(this bool? value, Translation translation)
   {
      if (value is null) { return string.Empty; }
      return value.Value ? translation.Yes : translation.No;
   }
}