using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders;

internal class OperationWorksheetBuilder(
   IXLWorkbook workbook,
   OpenApiDocumentationOptions options,
   ObjectLinkRegistry? objectLinks = null)
   : WorksheetBuilder(options)
{
   private readonly RowPointer _actualRowPointer = new(1);
   private IXLWorksheet _worksheet = null!;
   private int _attributesColumnsStartIndex;

   public IXLWorksheet Build(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation)
   {
      var worksheet = GetWorksheetName(path, operation, operationType);

      CreateNewWorksheet(worksheet);
      _actualRowPointer.GoTo(1);

      _attributesColumnsStartIndex = MaxPropertiesTreeLevel.Calculate(operation, Options.MaxDepth);
      PropertiesTreeColumns.NarrowTreeColumns(_worksheet, _attributesColumnsStartIndex);

      AddHomePageLink();
      AddOperationInfos(path, pathItem, operationType, operation);
      AddRequestParameters(operation);
      AddRequestBody(operation);
      AddResponseBody(operation);
      PropertiesTreeColumns.AdjustToContents(_worksheet, _attributesColumnsStartIndex);

      return _worksheet;
   }

   private string GetWorksheetName(string path, OpenApiOperation operation, OperationType operationType)
   {
      // Short enough to leave room for the number a repeated name gets, inside the 31 characters
      // Excel allows.
      const int maxLength = 28;

      string name;
      if (!string.IsNullOrEmpty(operation.OperationId))
      {
         // take worksheet name from OperationId
         name = operation.OperationId;
      }
      else
      {
         // generate worksheet name based on operationType and path
         var pathName = path.Replace("/", "-");
         name = operationType.ToString().ToUpper() + "_" + pathName[1..];
      }

      return WorksheetName.Unique(workbook, name, maxLength);
   }

   private void CreateNewWorksheet(string operation)
   {
      _worksheet = workbook.Worksheets.Add(operation);
      _worksheet.Style.Font.FontSize = 10;
      _worksheet.Style.Font.FontName = "Arial";
      _worksheet.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Left;
      _worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
   }

   private void AddOperationInfos(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation) =>
      new OperationInfoBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
         .AddOperationInfoSection(path, pathItem, operationType, operation);

   private void AddRequestParameters(OpenApiOperation operation) =>
      new RequestParametersBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options, objectLinks)
         .AddRequestParametersPart(operation);

   private void AddRequestBody(OpenApiOperation operation) =>
      new RequestBodyBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options, objectLinks)
         .AddRequestBodyPart(operation);

   private void AddResponseBody(OpenApiOperation operation) =>
      new ResponseBodyBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options, objectLinks)
         .AddResponseBodyPart(operation);

   private void AddHomePageLink() => new HomePageLinkBuilder(_actualRowPointer, _worksheet, Options)
      .AddHomePageLinkSection();
}