using Microsoft.OpenApi.Models;

namespace openapi2excel.core.Builders;

/// <summary>
/// One operation of the specification, together with the worksheet documenting it.
/// </summary>
internal sealed class DocumentedOperation(
   string path,
   OpenApiPathItem pathItem,
   OperationType operationType,
   OpenApiOperation operation,
   string worksheetName)
{
   public string Path { get; } = path;

   public OpenApiPathItem PathItem { get; } = pathItem;

   public OperationType OperationType { get; } = operationType;

   public OpenApiOperation Operation { get; } = operation;

   public string WorksheetName { get; } = worksheetName;

   /// <summary>
   /// Tags of the operation, in the order the specification lists them, without the ones carrying
   /// no name and without a tag repeated on the same operation.
   /// </summary>
   public IReadOnlyList<OpenApiTag> Tags { get; } = ReadTags(operation);

   /// <summary>
   /// What the operation is for, in one line: its summary, or the first line of its description
   /// when it has no summary. Empty when the specification says neither.
   /// </summary>
   public string Summary => FirstLine(Operation.Summary) is { Length: > 0 } summary
      ? summary
      : FirstLine(Operation.Description);

   private static IReadOnlyList<OpenApiTag> ReadTags(OpenApiOperation operation)
   {
      if (operation.Tags is null)
      {
         return [];
      }

      var tags = new List<OpenApiTag>();
      foreach (var tag in operation.Tags)
      {
         var name = OperationsIndex.NameOf(tag);
         if (name.Length > 0 && !tags.Any(added => OperationsIndex.NameOf(added).Equals(name, StringComparison.Ordinal)))
         {
            tags.Add(tag);
         }
      }

      return tags;
   }

   private static string FirstLine(string? text)
      => text?.Split('\n', '\r').FirstOrDefault(line => !string.IsNullOrWhiteSpace(line))?.Trim() ?? string.Empty;
}

/// <summary>
/// The operations of one path, inside one tag group.
/// </summary>
internal sealed class IndexedPath(string path, string? summary, IReadOnlyList<DocumentedOperation> operations)
{
   public string Path { get; } = path;

   /// <summary>Summary of the path item, as the specification states it. Empty when it states none.</summary>
   public string Summary { get; } = summary ?? string.Empty;

   public IReadOnlyList<DocumentedOperation> Operations { get; } = operations;
}

/// <summary>
/// The paths of one tag, the group the specification itself puts an operation in.
/// </summary>
internal sealed class IndexedTag(string name, string? description, IReadOnlyList<IndexedPath> paths)
{
   public string Name { get; } = name;

   /// <summary>Description of the tag, as the specification states it. Empty when it states none.</summary>
   public string Description { get; } = description ?? string.Empty;

   public IReadOnlyList<IndexedPath> Paths { get; } = paths;
}

/// <summary>
/// Groups the operations of a document the way the specification groups them: by tag, and inside a
/// tag by path. An operation carrying several tags belongs to each of them, the way a reader of the
/// specification finds it under each of them.
/// </summary>
internal static class OperationsIndex
{
   /// <summary>
   /// Builds the groups, in the order the document declares its tags. A tag used by an operation
   /// but never declared follows the declared ones, in the order the operations first use it, and
   /// the operations carrying no tag at all end up in a last group named <paramref name="untaggedGroupName"/>.
   /// </summary>
   public static IReadOnlyList<IndexedTag> Group(OpenApiDocument document,
      IEnumerable<DocumentedOperation> operations, string untaggedGroupName)
   {
      var operationsOfTag = new Dictionary<string, List<DocumentedOperation>>(StringComparer.Ordinal);
      var descriptionOfTag = new Dictionary<string, string>(StringComparer.Ordinal);
      var untagged = new List<DocumentedOperation>();
      var usedTags = new List<string>();

      foreach (var documented in operations)
      {
         if (documented.Tags.Count == 0)
         {
            untagged.Add(documented);
            continue;
         }

         foreach (var tag in documented.Tags)
         {
            var name = NameOf(tag);
            if (!operationsOfTag.TryGetValue(name, out var tagged))
            {
               operationsOfTag[name] = tagged = [];
               usedTags.Add(name);
            }

            tagged.Add(documented);
            RememberDescription(descriptionOfTag, name, tag.Description);
         }
      }

      // The document describes its tags in one place, the operations only point at them, so what
      // the document says about a tag wins over what a copy of it carries.
      foreach (var declared in DeclaredTags(document))
      {
         RememberDescription(descriptionOfTag, NameOf(declared), declared.Description, overwrite: true);
      }

      var groups = OrderTags(document, usedTags)
         .Select(name => new IndexedTag(name, Described(descriptionOfTag, name), GroupByPath(operationsOfTag[name])))
         .ToList();

      if (untagged.Count > 0)
      {
         groups.Add(new IndexedTag(untaggedGroupName, null, GroupByPath(untagged)));
      }

      return groups;
   }

   /// <summary>
   /// Name of a tag. A tag written as a reference to a declared one carries its name in the
   /// reference until the reader resolves it, so both places are read.
   /// </summary>
   public static string NameOf(OpenApiTag tag)
      => (string.IsNullOrWhiteSpace(tag.Name) ? tag.Reference?.Id : tag.Name)?.Trim() ?? string.Empty;

   private static IEnumerable<OpenApiTag> DeclaredTags(OpenApiDocument document)
      => document.Tags ?? Enumerable.Empty<OpenApiTag>();

   /// <summary>
   /// The tags actually used by an operation, ordered by the declaration of the document first,
   /// then by first use for the ones the document never declares.
   /// </summary>
   private static IEnumerable<string> OrderTags(OpenApiDocument document, IReadOnlyList<string> usedTags)
   {
      var used = new HashSet<string>(usedTags, StringComparer.Ordinal);
      var ordered = new List<string>();
      var taken = new HashSet<string>(StringComparer.Ordinal);

      foreach (var name in DeclaredTags(document).Select(NameOf))
      {
         if (used.Contains(name) && taken.Add(name))
         {
            ordered.Add(name);
         }
      }

      ordered.AddRange(usedTags.Where(name => taken.Add(name)));
      return ordered;
   }

   private static IReadOnlyList<IndexedPath> GroupByPath(IReadOnlyList<DocumentedOperation> operations)
   {
      var operationsOfPath = new Dictionary<string, List<DocumentedOperation>>(StringComparer.Ordinal);
      var paths = new List<string>();

      foreach (var operation in operations)
      {
         if (!operationsOfPath.TryGetValue(operation.Path, out var samePath))
         {
            operationsOfPath[operation.Path] = samePath = [];
            paths.Add(operation.Path);
         }

         samePath.Add(operation);
      }

      return paths
         .Select(path => new IndexedPath(path, operationsOfPath[path][0].PathItem.Summary, operationsOfPath[path]))
         .ToList();
   }

   private static void RememberDescription(IDictionary<string, string> descriptions, string name,
      string? description, bool overwrite = false)
   {
      if (string.IsNullOrWhiteSpace(description))
      {
         return;
      }

      if (overwrite || !descriptions.ContainsKey(name))
      {
         descriptions[name] = description!.Trim();
      }
   }

   private static string? Described(IReadOnlyDictionary<string, string> descriptions, string name)
      => descriptions.TryGetValue(name, out var description) ? description : null;
}
