using ClosedXML.Excel;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

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
}