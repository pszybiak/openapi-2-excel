using ClosedXML.Excel;

namespace openapi2excel.Builders.WorksheetPartsBuilders
{
    internal abstract class WorksheetPartBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
    {
        protected OpenApiDocumentationOptions Options { get; } = options;
        protected RowPointer ActualRow { get; } = actualRow;
        protected IXLWorksheet Worksheet { get; } = worksheet;

        protected void MoveToNextRow()
        {
            ActualRow.MoveNext();
        }

        protected void AddEmptyRow()
            => MoveToNextRow();

        protected void FillCell(int column, string? value, XLColor? backgoundColor = null)
        {
            Worksheet.Cell(ActualRow, column).Value = value;
            if (backgoundColor is not null)
            {
                Worksheet.Cell(ActualRow, column).Style.Fill.BackgroundColor = backgoundColor;
            }
        }

        protected void FillCell(int column, XLColor backgoundColor)
        {
            Worksheet.Cell(ActualRow, column).Style.Fill.BackgroundColor = backgoundColor;
        }
    }

    internal class RowPointer(int rowNumber)
    {
        public static implicit operator int(RowPointer d) => d.Get();
        public static explicit operator RowPointer(int b) => new(b);

        public void MoveNext(int rowCount = 1)
        {
            rowNumber += rowCount;
        }

        public int Get() => rowNumber;

        public int GoTo(int row)
        {
            return rowNumber = row;
        }
    }
}
