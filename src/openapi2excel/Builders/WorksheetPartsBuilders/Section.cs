using ClosedXML.Excel;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class Section(IXLWorksheet worksheet, RowPointer actualRow) : IDisposable
{
   private readonly RowPointer _startRow = actualRow.Copy();

   public void Dispose()
   {
      worksheet.Rows(_startRow.Get(), actualRow.Get()).Group();
   }
}