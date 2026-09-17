using ClosedXML.Excel;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

/// <summary>
/// Width of the columns of the table describing a schema, the same table on every worksheet built
/// out of one: an operation, and an object of the objects worksheet.
/// </summary>
internal static class PropertiesTreeColumns
{
   /// <summary>
   /// The columns the tree uses to indent a nested property hold no text of their own, they only
   /// have to be wide enough to be seen.
   /// </summary>
   public static void NarrowTreeColumns(IXLWorksheet worksheet, int attributesColumnsStartIndex)
   {
      for (var columnIndex = 1; columnIndex < attributesColumnsStartIndex - 1; columnIndex++)
      {
         worksheet.Column(columnIndex).Width = 1.8;
      }
   }

   /// <summary>
   /// Widens the columns a name, a type or a description would otherwise be cut off in: the last of
   /// the columns the tree indents a name into, the name of the referenced type, and the
   /// description at the end of the row.
   /// </summary>
   public static void AdjustToContents(IXLWorksheet worksheet, int attributesColumnsStartIndex)
   {
      if (attributesColumnsStartIndex > 1)
      {
         worksheet.Column(attributesColumnsStartIndex - 1).AdjustToContents();
      }

      // A referenced type name is a fully qualified name in a generated specification, too long for
      // the default column width and too long to let it push the rest of the row off the screen.
      const int objectTypeColumnOffset = 3;
      const double minObjectTypeWidth = 13;
      const double maxObjectTypeWidth = 40;
      worksheet.Column(attributesColumnsStartIndex + objectTypeColumnOffset)
         .AdjustToContents(minObjectTypeWidth, maxObjectTypeWidth);

      worksheet.LastColumnUsed()?.AdjustToContents();
   }
}
