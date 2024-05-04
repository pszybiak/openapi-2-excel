using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
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

      protected CellBuilder Fill(int column)
         => new(Worksheet.Cell(ActualRow, column), Options);

      protected void AddPropertiesTreeForMediaTypes(IDictionary<string, OpenApiMediaType> mediaTypes,
         int attributesColumnIndex)
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
         foreach (var mediaType in mediaTypes)
         {
            var bodyFormatRowPointer = ActualRow.Copy();
            Cell(1).SetTextBold($"Body format: {mediaType.Key}");
            ActualRow.MoveNext();

            var columnCount = builder.AddPropertiesTree(ActualRow, mediaType.Value.Schema);
            Worksheet.Cell(bodyFormatRowPointer, 1).SetBackground(columnCount, HeaderBackgroundColor);
            ActualRow.MoveNext();
         }
      }
   }
}