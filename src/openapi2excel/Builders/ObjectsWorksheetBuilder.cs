using ClosedXML.Excel;
using openapi2excel.core.Builders.WorksheetPartsBuilders;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders;

/// <summary>
/// The last worksheet of the document: every object the operations use, one section each, described
/// by the same table the request and the response bodies are described with. The Object type column
/// of every worksheet links here, and this worksheet links back to the first one.
/// </summary>
internal class ObjectsWorksheetBuilder(
   IXLWorkbook workbook,
   OpenApiDocumentationOptions options,
   ObjectLinkRegistry objectLinks)
   : WorksheetBuilder(options)
{
   private const int WorksheetNameMaxLength = 31;

   private readonly RowPointer _actualRowPointer = new(1);
   private IXLWorksheet _worksheet = null!;
   private int _attributesColumnsStartIndex;

   /// <summary>
   /// Builds the worksheet, and tells the registry where every section starts. A specification
   /// whose operations name no object at all gets no worksheet, the way it gets no index when it
   /// declares no operation.
   /// </summary>
   public IXLWorksheet? Build(IReadOnlyList<DocumentedObject> objects)
   {
      if (objects.Count == 0)
      {
         return null;
      }

      CreateNewWorksheet();
      _actualRowPointer.GoTo(1);

      _attributesColumnsStartIndex =
         MaxPropertiesTreeLevel.Calculate(objects.Select(documented => documented.Schema), Options.MaxDepth);
      PropertiesTreeColumns.NarrowTreeColumns(_worksheet, _attributesColumnsStartIndex);

      AddHomePageLink();
      AddHeader();

      // Nothing groups the sections together: the name of an object is what a reader looks one up
      // by, so no collapsing may hide it. Each section collapses from the row naming its columns,
      // and collapsing every one of them leaves the list of the names.
      objects.ForEach(AddObject);

      PropertiesTreeColumns.AdjustToContents(_worksheet, _attributesColumnsStartIndex);
      return _worksheet;
   }

   private void CreateNewWorksheet()
   {
      _worksheet = workbook.Worksheets.Add(
         WorksheetName.Unique(workbook, Options.Translation.ObjectsSheetName, WorksheetNameMaxLength));
      _worksheet.Style.Font.FontSize = 10;
      _worksheet.Style.Font.FontName = "Arial";
      _worksheet.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Left;
      _worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
   }

   private void AddHomePageLink()
      => new HomePageLinkBuilder(_actualRowPointer, _worksheet, Options).AddHomePageLinkSection();

   private void AddHeader()
   {
      _worksheet.Cell(_actualRowPointer, 1).SetTextBold(Options.Translation.ObjectsHeader);
      _actualRowPointer.MoveNext();
   }

   /// <summary>
   /// One object: its name over the table describing it, collapsing into that name, and the row of
   /// the name is what every Object type cell naming the object links to.
   /// </summary>
   private void AddObject(DocumentedObject documented)
   {
      var titleRow = _actualRowPointer.Get();
      objectLinks.AddAnchor(documented.Name, _worksheet, titleRow);

      var lastColumn = new PropertiesTreeBuilder(_attributesColumnsStartIndex, _worksheet, Options, objectLinks)
         .AddTitledPropertiesTree(_actualRowPointer, documented.Name, documented.Schema, Options,
            titleNamesTheSchema: true);

      AddDescription(documented, titleRow, lastColumn);
   }

   /// <summary>
   /// What the specification says about the object, next to its name, in the column every
   /// description of the table below it already uses.
   /// </summary>
   private void AddDescription(DocumentedObject documented, int titleRow, int descriptionColumn)
   {
      var description = documented.Schema.Description.StripHtmlTags();
      if (!string.IsNullOrWhiteSpace(description))
      {
         _worksheet.Cell(titleRow, descriptionColumn).SetText(description);
      }
   }
}
