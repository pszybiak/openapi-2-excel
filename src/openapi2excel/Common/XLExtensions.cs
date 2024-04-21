using ClosedXML.Excel;

namespace openapi2excel.core.Common;

// ReSharper disable once InconsistentNaming
internal static class XLExtensions
{
   public static IXLCell SetBoldStyle(this IXLCell cell)
   {
      cell.Style.Font.SetBold(true);
      return cell;
   }

   public static IXLCell SetText(this IXLCell cell, string value)
      => cell.SetValue(value);

   public static IXLCell SetTextBold(this IXLCell cell, string value)
      => cell.SetBoldStyle().SetValue(value);

   public static IXLCell NextRow(this IXLCell cell)
      => cell.Worksheet.Cell(cell.Address.RowNumber + 1, 1);

   public static IXLCell If(this IXLCell cell, bool condition, Func<IXLCell, IXLCell> func)
   {
      return condition ? func(cell) : cell;
   }

   public static IXLCell IfNotEmpty(this IXLCell cell, string text, Func<IXLCell, IXLCell> func)
   {
      return string.IsNullOrEmpty(text) ? cell : func(cell);
   }
}