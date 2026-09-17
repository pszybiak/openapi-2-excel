using ClosedXML.Excel;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders;

/// <summary>
/// Turns every cell naming a schema, a value of the Object type column or the title of a body
/// format, into a link to the section of the objects worksheet documenting that schema.
/// <para>
/// A cell is written long before the section it points at exists, the objects worksheet is the last
/// one built, so the cells are collected while the document is written and linked at the end.
/// </para>
/// </summary>
internal sealed class ObjectLinkRegistry
{
   private readonly List<(IXLCell Cell, string ObjectName)> _links = [];
   private readonly Dictionary<string, string> _anchors = new(StringComparer.Ordinal);

   /// <summary>Remembers a cell naming an object, to be linked by <see cref="Resolve"/>.</summary>
   public void Register(IXLCell cell, string objectName) => _links.Add((cell, objectName));

   /// <summary>Remembers where the section documenting an object starts.</summary>
   public void AddAnchor(string objectName, IXLWorksheet worksheet, int rowNumber)
      => _anchors[objectName] = worksheet.AddressOfCell(rowNumber);

   /// <summary>
   /// Links every remembered cell. A cell naming a schema with no section of its own, which the
   /// catalog and the worksheet together are meant to make impossible, keeps the name it already
   /// holds instead of pointing nowhere.
   /// </summary>
   public void Resolve()
   {
      foreach (var (cell, objectName) in _links)
      {
         if (!_anchors.TryGetValue(objectName, out var address))
         {
            continue;
         }

         cell.SetHyperlink(new XLHyperlink(address));
         cell.SetLinkStyle();
      }
   }
}
