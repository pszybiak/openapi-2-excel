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