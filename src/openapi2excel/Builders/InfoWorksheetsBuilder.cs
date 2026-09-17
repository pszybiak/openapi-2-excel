using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders;

internal class InfoWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
   : WorksheetBuilder(options)
{
   private const int TagColumn = 1;
   private const int PathColumn = 2;

   /// <summary>Holds the type of an operation and the id the specification names it with.</summary>
   private const int OperationColumn = 3;

   private const int SummaryColumn = 4;
   private const int LastColumn = SummaryColumn;

   private static XLColor HeaderBackgroundColor => XLColor.LightGray;
   private static XLColor TagBackgroundColor => XLColor.FromArgb(0xE8, 0xE8, 0xE8);

   private readonly List<DocumentedOperation> _operations = [];
   private OpenApiDocument _readResultOpenApiDocument = null!;
   private IXLWorksheet _worksheet = null!;
   private int _actualRowIndex = 1;

   public IXLWorksheet Build(OpenApiDocument openApiDocument)
   {
      _readResultOpenApiDocument = openApiDocument;
      _worksheet = workbook.Worksheets.Add(Options.Translation.InfoSheetName);
      _worksheet.Style.Font.FontSize = 10;
      _worksheet.Style.Font.FontName = "Arial";

      // The index below groups tag over path over operation, and the reader collapses a group from
      // the header above it, the way every operation worksheet already behaves.
      _worksheet.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Left;
      _worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;

      // The tag and the path only have to carry the indentation of the group, their text spills
      // over the empty cells of the row to the right of them.
      _worksheet.Column(TagColumn).Width = WidthOfPixels(30);
      _worksheet.Column(PathColumn).Width = WidthOfPixels(40);
      _worksheet.Column(OperationColumn).Width = 15;
      _worksheet.Column(SummaryColumn).Width = 70;

      AddVersion();
      AddTitle();
      AddDescription();

      return _worksheet;
   }

   /// <summary>
   /// Remembers an operation and the worksheet documenting it. Nothing is written before
   /// <see cref="AddOperationsIndex"/>, because an operation belongs to a group the last operation
   /// of the document can still join.
   /// </summary>
   public void AddLink(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation, IXLWorksheet worksheet)
      => _operations.Add(new DocumentedOperation(path, pathItem, operationType, operation, worksheet.Name));

   /// <summary>
   /// Writes the index of the operations, grouped the way the specification groups them: by tag,
   /// and inside a tag by path.
   /// </summary>
   public void AddOperationsIndex()
   {
      if (_operations.Count == 0)
      {
         return;
      }

      _actualRowIndex++;
      AddIndexHeader();

      var firstIndexRow = _actualRowIndex;
      foreach (var tag in OperationsIndex.Group(_readResultOpenApiDocument, _operations,
                  Options.Translation.InfoUntagged))
      {
         AddTagGroup(tag);
      }

      // An operationId of a generated specification is a long name, and the column holding it is
      // measured on the index alone: the block describing the api names its version and its title
      // in that column too, and neither is as long.
      const double minOperationWidth = 15;
      const double maxOperationWidth = 45;
      _worksheet.Column(OperationColumn)
         .AdjustToContents(firstIndexRow, _actualRowIndex - 1, minOperationWidth, maxOperationWidth);
   }

   /// <summary>
   /// Writes the link to the worksheet documenting the objects, under the index of the operations.
   /// That worksheet links back to this one, and nothing else on this one would tell a reader that
   /// it exists at all.
   /// </summary>
   public void AddObjectsLink(IXLWorksheet objectsWorksheet)
   {
      _actualRowIndex++;

      // The label spills over the columns to its right, which are as narrow as the indentation of a
      // group needs, so the whole width of the index is what a reader clicks.
      for (var column = TagColumn; column <= LastColumn; column++)
      {
         var cell = _worksheet.Cell(_actualRowIndex, column);
         cell.SetHyperlink(new XLHyperlink(objectsWorksheet.AddressOfCell()));
         cell.SetLinkStyle();
      }

      _worksheet.Cell(_actualRowIndex, TagColumn).SetTextBold(Options.Translation.ObjectsHeader);
      _actualRowIndex++;
   }

   private void AddIndexHeader()
   {
      _worksheet.Cell(_actualRowIndex++, TagColumn).SetTextBold(Options.Translation.InfoOperations);

      _worksheet.Cell(_actualRowIndex, TagColumn).SetTextBold(Options.Translation.InfoTag);
      _worksheet.Cell(_actualRowIndex, PathColumn).SetTextBold(Options.Translation.OperationPath);
      _worksheet.Cell(_actualRowIndex, OperationColumn).SetTextBold(string.Format(
         Options.Translation.InfoOperationWithId,
         Options.Translation.OperationType,
         Options.Translation.OperationId));
      _worksheet.Cell(_actualRowIndex, SummaryColumn).SetTextBold(Options.Translation.OperationSummary);
      _worksheet.Cell(_actualRowIndex, TagColumn)
         .SetBackground(LastColumn, HeaderBackgroundColor)
         .SetBottomBorder(LastColumn);
      _actualRowIndex++;
   }

   private void AddTagGroup(IndexedTag tag)
   {
      var headerRowIndex = _actualRowIndex;
      _worksheet.Cell(headerRowIndex, TagColumn).SetTextBold(tag.Name).SetBackground(LastColumn, TagBackgroundColor);
      SetTextIfAny(headerRowIndex, SummaryColumn, tag.Description);
      _actualRowIndex++;

      tag.Paths.ForEach(AddPathGroup);

      GroupRows(headerRowIndex + 1, _actualRowIndex - 1);
   }

   private void AddPathGroup(IndexedPath path)
   {
      var headerRowIndex = _actualRowIndex;
      _worksheet.Cell(headerRowIndex, PathColumn).SetTextBold(path.Path);
      SetTextIfAny(headerRowIndex, SummaryColumn, path.Summary);
      _actualRowIndex++;

      path.Operations.ForEach(AddOperation);

      GroupRows(headerRowIndex + 1, _actualRowIndex - 1);
   }

   private void AddOperation(DocumentedOperation operation)
   {
      // The id is what the specification calls the operation and what names the worksheet
      // documenting it. An operation stating none is named after its method and its path instead,
      // and the type of the operation is all this cell has to say.
      var operationType = operation.OperationType.ToString().ToUpper();
      var operationId = operation.Operation.OperationId;
      _worksheet.Cell(_actualRowIndex, OperationColumn)
         .SetText(string.IsNullOrEmpty(operationId)
            ? operationType
            : string.Format(Options.Translation.InfoOperationWithId, operationType, operationId))
         .SetHyperlink(LinkTo(operation));

      var summary = operation.Summary;
      if (!string.IsNullOrEmpty(summary))
      {
         _worksheet.Cell(_actualRowIndex, SummaryColumn)
            .SetText(summary)
            .SetHyperlink(LinkTo(operation));
      }

      _actualRowIndex++;
   }

   private static XLHyperlink LinkTo(DocumentedOperation operation)
      => new(XLExtensions.AddressOfCell(operation.WorksheetName));

   /// <summary>
   /// A column width, in pixels. Excel measures a column in characters of the default font of the
   /// workbook, Calibri 11, whose widest digit takes 7 pixels, plus 5 pixels of padding.
   /// </summary>
   private static double WidthOfPixels(double pixels) => (pixels - 5) / 7;

   private void SetTextIfAny(int rowIndex, int columnIndex, string text)
   {
      if (!string.IsNullOrWhiteSpace(text))
      {
         _worksheet.Cell(rowIndex, columnIndex).SetText(text);
      }
   }

   /// <summary>
   /// Collapses the rows of one group under the header row above them. An empty group, a tag with
   /// no path or a path with no operation, has nothing to collapse.
   /// </summary>
   private void GroupRows(int firstRowIndex, int lastRowIndex)
   {
      if (lastRowIndex >= firstRowIndex)
      {
         _worksheet.Rows(firstRowIndex, lastRowIndex).Group();
      }
   }

   private void AddVersion() => FillInfo(Options.Translation.InfoVersion, _readResultOpenApiDocument.Info.Version);

   private void AddDescription() => FillInfo(Options.Translation.InfoDescription, _readResultOpenApiDocument.Info.Description, true);

   private void AddTitle() => FillInfo(Options.Translation.InfoTitle, _readResultOpenApiDocument.Info.Title);

   /// <summary>
   /// Writes one line of the block describing the api, in the two columns the index keeps for
   /// the name of a thing and for its text. The two columns to the left of them are only as wide
   /// as the indentation of a group needs, too narrow to hold a label.
   /// </summary>
   private void FillInfo(string name, string value, bool splitMultipleRowText = false)
   {
      if (string.IsNullOrEmpty(value))
      {
         return;
      }

      _worksheet.Cell(_actualRowIndex, OperationColumn).SetTextBold(name);
      if (splitMultipleRowText)
      {
         value.Split('\n', '\r', StringSplitOptions.RemoveEmptyEntries)
            .Select(v => v.Trim())
            .Where(v => !string.IsNullOrEmpty(v))
            .ForEach(splittedValue => _worksheet.Cell(_actualRowIndex++, SummaryColumn).Value = splittedValue);
      }
      else
      {
         _worksheet.Cell(_actualRowIndex++, SummaryColumn).Value = value;
      }
   }
}
