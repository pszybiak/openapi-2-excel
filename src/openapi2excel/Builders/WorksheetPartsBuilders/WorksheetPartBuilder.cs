using ClosedXML.Excel;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders
{
   internal abstract class WorksheetPartBuilder(
      RowPointer actualRow,
      IXLWorksheet worksheet,
      OpenApiDocumentationOptions options)
   {
      protected OpenApiDocumentationOptions Options { get; } = options;
      protected RowPointer ActualRow { get; } = actualRow;
      protected IXLWorksheet Worksheet { get; } = worksheet;
      protected XLColor HeaderBackgroundColor => XLColor.LightGray;

      protected IXLCell Cell(int column)
         => Worksheet.Cell(ActualRow.Get(), column);
   }
}