using ClosedXML.Excel;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class HomePageLinkBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddHomePageLinkSection()
   {
      Cell(1).SetValue("<<<<<")
         .SetHyperlink(new XLHyperlink(XLExtensions.AddressOfCell(Options.Translation.InfoSheetName)));
      ActualRow.MoveNext(2);
   }
}