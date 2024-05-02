using ClosedXML.Excel;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

[System.Diagnostics.CodeAnalysis.SuppressMessage("Major Code Smell", "S3881:\"IDisposable\" should be implemented correctly", Justification = "<Pending>")]
internal class Section(IXLWorksheet worksheet, RowPointer actualRow) : IDisposable
{
   private readonly RowPointer _startRow = actualRow.Copy();

   public void Dispose()
   {
      worksheet.Rows(_startRow.Get(), actualRow.Get()).Group();
   }
}