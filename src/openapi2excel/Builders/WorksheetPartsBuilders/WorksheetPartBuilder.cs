using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using OpenApi2Excel.Common;

namespace OpenApi2Excel.Builders.WorksheetPartsBuilders
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

        protected void AddEmptyRow()
            => ActualRow.MoveNext();

        protected CellBuilder Fill(int column)
            => new(Worksheet.Cell(ActualRow, column), Options);

        protected CellBuilder FillHeader(int column)
            => Fill(column).WithBackground(HeaderBackgroundColor);

        protected void AddPropertiesTreeForMediaTypes(IDictionary<string, OpenApiMediaType> mediaTypes, int attributesColumnIndex)
        {
            var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
            foreach (var mediaType in mediaTypes)
            {
                AddMediaTypeFormat(mediaType.Key);
                builder.AddPropertiesTree(ActualRow, mediaType.Value.Schema);
                AddEmptyRow();
            }
            AddEmptyRow();

            void AddMediaTypeFormat(string name)
            {
                Fill(1).WithText($"Format: {name}").WithBoldStyle();
                ActualRow.MoveNext();
            }
        }

        protected void AddProperty(string name, OpenApiSchema schema, int level, int attributesColumnIndex)
        {
            // current property
            AddPropertyRow(name, schema, level++);
            AddProperties(schema, level, attributesColumnIndex);

            return;

            void AddPropertyRow(string propertyName, OpenApiSchema propertySchema, int propertyLevel)
            {
                FillHeaderBackground(1, propertyLevel - 1);
                FillCell(propertyLevel, propertyName);
                FillSchemaDescriptionCells(propertySchema, attributesColumnIndex);
                ActualRow.MoveNext();
            }
        }

        protected void AddProperties(OpenApiSchema schema, int level, int attributesColumnIndex)
        {
            if (schema.Items != null)
            {
                if (schema.Items.Properties.Any())
                {
                    // array of object properties
                    foreach (var property in schema.Items.Properties)
                    {
                        AddProperty(property.Key, property.Value, level, attributesColumnIndex);
                    }
                }
                else
                {
                    // if array contains simple type items
                    AddProperty("element", schema.Items, level, attributesColumnIndex);
                }
            }

            // subproperties
            foreach (var property in schema.Properties)
            {
                AddProperty(property.Key, property.Value, level, attributesColumnIndex);
            }
        }

        protected void FillSchemaDescriptionCells(OpenApiSchema schema, int startColumn)
        {
            Fill(startColumn).WithText(schema.Type)
                .Next().WithText(schema.Format)
                .Next().WithText(schema.GetPropertyLengthDescription()).WithHorizontalAlignment(XLAlignmentHorizontalValues.Center)
                .Next().WithText(schema.GetPropertyRangeDescription()).WithHorizontalAlignment(XLAlignmentHorizontalValues.Center)
                .Next().WithText(schema.Pattern)
                .Next().WithText(schema.Description);
        }

        protected int FillSchemaDescriptionHeaderCells(int attributesStartColumn)
        {
            var cellBuilder = Fill(1).WithText("Name").WithBoldStyle()
                .Next(attributesStartColumn - 1).WithText("Type").WithBoldStyle()
                .Next().WithText("Format").WithBoldStyle()
                .Next().WithText("Length").WithBoldStyle()
                .Next().WithText("Range").WithBoldStyle()
                .Next().WithText("Pattern").WithBoldStyle()
                .Next().WithText("Description").WithBoldStyle();

            var lastUsedColumn = cellBuilder.GetCell().Address.ColumnNumber;
            AddBottomBorder(1, lastUsedColumn);
            FillHeaderBackground(1, lastUsedColumn);
            return lastUsedColumn;
        }

        protected void FillCell(int column, string? value, XLColor? backgoundColor = null, XLAlignmentHorizontalValues alignment = XLAlignmentHorizontalValues.Left)
        {
            var cellBuilder = Fill(column).WithText(value);
            if (backgoundColor is not null)
            {
                cellBuilder.WithBackground(backgoundColor);
            }
            cellBuilder.WithHorizontalAlignment(alignment);
        }

        protected void FillBackground(int startColumn, int endColumn, XLColor backgroundColor)
        {
            var cellBuilder = Fill(startColumn);
            for (var columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
            {
                cellBuilder.WithBackground(backgroundColor).Next();
            }
        }

        protected void FillHeaderBackground(int startColumn, int endColumn)
            => FillBackground(startColumn, endColumn, HeaderBackgroundColor);

        protected void AddBottomBorder(int startColumn, int endColumn)
        {
            var cellBuilder = Fill(startColumn);
            for (var columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
            {
                cellBuilder.WithBottomBorder().Next();
            }
        }
    }
}
