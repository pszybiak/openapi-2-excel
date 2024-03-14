using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using System.Text;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders
{
    internal abstract class WorksheetPartBuilder(RowPointer actualRow, IXLWorksheet worksheet, OpenApiDocumentationOptions options)
    {
        protected OpenApiDocumentationOptions Options { get; } = options;
        protected RowPointer ActualRow { get; } = actualRow;
        protected IXLWorksheet Worksheet { get; } = worksheet;

        protected void MoveToNextRow()
            => ActualRow.MoveNext();

        protected void AddEmptyRow()
            => MoveToNextRow();

        protected void FillSchemaDescriptionCells(OpenApiSchema schema, int startColumn)
        {
            FillCell(startColumn++, schema.Type);
            FillCell(startColumn++, schema.Format);
            FillCell(startColumn++, GetPropertyLength(schema));
            FillCell(startColumn++, GetPropertyRange(schema));
            FillCell(startColumn++, schema.Pattern);
            FillCell(startColumn, schema.Description);
        }

        protected void FillSchemaDescriptionHeaderCells(int startColumn)
        {
            FillHeaderCell("Type", startColumn++);
            FillHeaderCell("Format", startColumn++);
            FillHeaderCell("Length", startColumn++);
            FillHeaderCell("Range", startColumn++);
            FillHeaderCell("Pattern", startColumn++);
            FillHeaderCell("Description", startColumn);
        }

        protected void FillHeaderCell(int columnIndex)
            => FillCell(columnIndex, XLColor.LightBlue);

        protected void FillHeaderCell(string? label, int columnIndex)
            => FillCell(columnIndex, label, XLColor.LightBlue);

        protected void FillCell(int column, string? value, XLColor? backgoundColor = null)
        {
            Worksheet.Cell(ActualRow, column).Value = value;
            if (backgoundColor is not null)
            {
                Worksheet.Cell(ActualRow, column).Style.Fill.BackgroundColor = backgoundColor;
            }
        }

        protected void FillCell(int column, bool value, XLColor? backgoundColor = null)
        {
            Worksheet.Cell(ActualRow, column).Value = Options.Language.Get(value);
            if (backgoundColor is not null)
            {
                Worksheet.Cell(ActualRow, column).Style.Fill.BackgroundColor = backgoundColor;
            }
        }

        protected void FillCell(int column, XLColor backgoundColor)
        {
            Worksheet.Cell(ActualRow, column).Style.Fill.BackgroundColor = backgoundColor;
        }

        private static string GetPropertyLength(OpenApiSchema schema)
        {
            StringBuilder propertyTypeDescription = new();
            if (schema.MinLength is not null)
            {
                propertyTypeDescription.Append($"{schema.MinLength}");
            }
            if (schema.MinLength is not null && !(schema.MinLength is null && schema.MaxLength is not null))
            {
                propertyTypeDescription.Append("..");
            }
            if (schema.MaxLength is not null)
            {
                propertyTypeDescription.Append(schema.MaxLength);
            }
            return propertyTypeDescription.ToString();
        }

        private static string GetPropertyRange(OpenApiSchema schema)
        {
            StringBuilder propertyTypeDescription = new();
            if (schema.Minimum is not null)
            {
                var sign = schema.ExclusiveMinimum is null or false ? "[" : "(";
                propertyTypeDescription.Append($"{sign}{schema.Minimum}");
            }
            else if (schema.Maximum is not null)
            {
                propertyTypeDescription.Append("(..");
            }
            if (schema.Minimum is not null || schema.Maximum is not null)
            {
                propertyTypeDescription.Append(';');
            }
            if (schema.Maximum is not null)
            {
                var sign = schema.ExclusiveMaximum is null or false ? "]" : ")";
                propertyTypeDescription.Append($"{schema.Maximum}{sign}");
            }
            else if (schema.Minimum is not null)
            {
                propertyTypeDescription.Append("..)");
            }
            return propertyTypeDescription.ToString();
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
