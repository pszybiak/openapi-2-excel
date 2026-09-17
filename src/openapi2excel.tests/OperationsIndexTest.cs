using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// Gate for the grouping behind the index of the info worksheet. The order of the groups, and
   /// what lands in which group, is decided here and only rendered by the worksheet builder, so it
   /// is asserted here instead of through cell addresses.
   /// </summary>
   public class OperationsIndexTest
   {
      private const string Untagged = "Other operations";

      [Fact]
      public void Operations_are_grouped_by_tag_then_by_path_keeping_the_order_of_the_document()
      {
         var document = Document(
            tags: [Tag("pet", "Everything about your Pets"), Tag("store", "Access to Petstore orders")],
            paths:
            [
               Path("/pet", "Pets", (OperationType.Put, "updatePet", ["pet"]), (OperationType.Post, "addPet", ["pet"])),
               Path("/store/order", null, (OperationType.Post, "placeOrder", ["store"])),
               Path("/pet/{petId}", null, (OperationType.Get, "getPetById", ["pet"]))
            ]);

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal(["pet", "store"], groups.Select(group => group.Name));
         Assert.Equal("Everything about your Pets", groups[0].Description);
         Assert.Equal(["/pet", "/pet/{petId}"], groups[0].Paths.Select(path => path.Path));
         Assert.Equal("Pets", groups[0].Paths[0].Summary);
         Assert.Equal(["updatePet", "addPet"], groups[0].Paths[0].Operations.Select(o => o.Operation.OperationId));
         Assert.Equal([OperationType.Put, OperationType.Post],
            groups[0].Paths[0].Operations.Select(o => o.OperationType));
         Assert.Equal(["/store/order"], groups[1].Paths.Select(path => path.Path));
      }

      [Fact]
      public void Groups_follow_the_order_the_document_declares_its_tags_not_the_order_the_paths_use_them()
      {
         var document = Document(
            tags: [Tag("user"), Tag("pet"), Tag("store")],
            paths:
            [
               Path("/store/order", null, (OperationType.Post, "placeOrder", ["store"])),
               Path("/pet", null, (OperationType.Put, "updatePet", ["pet"])),
               Path("/user", null, (OperationType.Post, "createUser", ["user"]))
            ]);

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal(["user", "pet", "store"], groups.Select(group => group.Name));
      }

      [Fact]
      public void Tag_the_document_declares_but_no_operation_uses_is_not_a_group()
      {
         var document = Document(
            tags: [Tag("pet"), Tag("deprecated-and-unused")],
            paths: [Path("/pet", null, (OperationType.Put, "updatePet", ["pet"]))]);

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal(["pet"], groups.Select(group => group.Name));
      }

      [Fact]
      public void Tag_no_declaration_describes_still_groups_after_the_declared_ones_in_order_of_first_use()
      {
         var document = Document(
            tags: [Tag("pet")],
            paths:
            [
               Path("/store/order", null, (OperationType.Post, "placeOrder", ["store"])),
               Path("/pet", null, (OperationType.Put, "updatePet", ["pet"])),
               Path("/admin", null, (OperationType.Get, "getAdmin", ["admin"]))
            ]);

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal(["pet", "store", "admin"], groups.Select(group => group.Name));
         Assert.Equal(string.Empty, groups[1].Description);
      }

      [Fact]
      public void Operation_carrying_several_tags_is_listed_under_each_of_them()
      {
         var document = Document(
            tags: [Tag("pet"), Tag("store")],
            paths: [Path("/pet/{petId}", null, (OperationType.Get, "getPetById", ["pet", "store"]))]);

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal(["pet", "store"], groups.Select(group => group.Name));
         Assert.Equal("getPetById", groups[0].Paths[0].Operations.Single().Operation.OperationId);
         Assert.Equal("getPetById", groups[1].Paths[0].Operations.Single().Operation.OperationId);
      }

      [Fact]
      public void Same_tag_repeated_on_one_operation_lists_it_once()
      {
         var document = Document(
            tags: [Tag("pet")],
            paths: [Path("/pet", null, (OperationType.Put, "updatePet", ["pet", "pet"]))]);

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Single(groups);
         Assert.Single(groups[0].Paths[0].Operations);
      }

      [Fact]
      public void Operations_carrying_no_tag_end_up_in_a_last_group_of_their_own()
      {
         var document = Document(
            tags: [Tag("pet")],
            paths:
            [
               Path("/health", null, (OperationType.Get, "health", [])),
               Path("/pet", null, (OperationType.Put, "updatePet", ["pet"]))
            ]);

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal(["pet", Untagged], groups.Select(group => group.Name));
         Assert.Equal("health", groups[1].Paths.Single().Operations.Single().Operation.OperationId);
      }

      [Fact]
      public void Document_with_no_tag_at_all_is_one_group()
      {
         var document = Document(
            tags: [],
            paths:
            [
               Path("/health", null, (OperationType.Get, "health", [])),
               Path("/ready", null, (OperationType.Get, "ready", []))
            ]);

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal([Untagged], groups.Select(group => group.Name));
         Assert.Equal(["/health", "/ready"], groups[0].Paths.Select(path => path.Path));
      }

      [Fact]
      public void What_the_document_says_about_a_tag_wins_over_the_copy_the_operation_carries()
      {
         var document = Document(
            tags: [Tag("pet", "Everything about your Pets")],
            paths: [Path("/pet", null, (OperationType.Put, "updatePet", ["pet"]))]);
         document.Paths["/pet"].Operations[OperationType.Put].Tags[0].Description = "A stale copy";

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal("Everything about your Pets", groups[0].Description);
      }

      [Fact]
      public void Description_of_a_tag_no_declaration_describes_is_taken_from_the_operation()
      {
         var document = Document(tags: [], paths: [Path("/pet", null, (OperationType.Put, "updatePet", ["pet"]))]);
         document.Paths["/pet"].Operations[OperationType.Put].Tags[0].Description = "Everything about your Pets";

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal("Everything about your Pets", groups[0].Description);
      }

      [Fact]
      public void Tag_left_as_an_unresolved_reference_groups_under_the_name_the_reference_carries()
      {
         // Until the reader resolves it, a tag written as a reference to a declared one carries its
         // name in the reference id and nowhere else.
         var document = Document(tags: [Tag("pet")], paths: [Path("/pet", null, (OperationType.Put, "updatePet", []))]);
         document.Paths["/pet"].Operations[OperationType.Put].Tags =
         [
            new OpenApiTag
            {
               UnresolvedReference = true,
               Reference = new OpenApiReference { Type = ReferenceType.Tag, Id = "pet" }
            }
         ];

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal(["pet"], groups.Select(group => group.Name));
      }

      [Fact]
      public void Tag_carrying_no_name_at_all_leaves_the_operation_untagged()
      {
         var document = Document(tags: [], paths: [Path("/pet", null, (OperationType.Put, "updatePet", []))]);
         document.Paths["/pet"].Operations[OperationType.Put].Tags = [new OpenApiTag { Name = "   " }];

         var groups = OperationsIndex.Group(document, Operations(document), Untagged);

         Assert.Equal([Untagged], groups.Select(group => group.Name));
      }

      [Fact]
      public void Summary_of_an_operation_falls_back_to_the_first_line_of_its_description()
      {
         var document = Document(tags: [], paths: [Path("/pet", null, (OperationType.Put, "updatePet", []))]);
         var operation = document.Paths["/pet"].Operations[OperationType.Put];
         operation.Summary = null;
         operation.Description = "  Update an existing pet by Id\nand return it  ";

         var documented = Operations(document).Single();

         Assert.Equal("Update an existing pet by Id", documented.Summary);
      }

      [Fact]
      public void Operation_stating_neither_summary_nor_description_has_no_summary()
      {
         var document = Document(tags: [], paths: [Path("/pet", null, (OperationType.Put, "updatePet", []))]);
         document.Paths["/pet"].Operations[OperationType.Put].Summary = null;

         Assert.Equal(string.Empty, Operations(document).Single().Summary);
      }

      private static List<DocumentedOperation> Operations(OpenApiDocument document)
         => document.Paths
            .SelectMany(path => path.Value.Operations.Select(operation => new DocumentedOperation(
               path.Key, path.Value, operation.Key, operation.Value, operation.Value.OperationId!)))
            .ToList();

      private static OpenApiTag Tag(string name, string? description = null)
         => new() { Name = name, Description = description };

      private static KeyValuePair<string, OpenApiPathItem> Path(string path, string? summary,
         params (OperationType Type, string OperationId, string[] Tags)[] operations)
      {
         var pathItem = new OpenApiPathItem { Summary = summary };
         foreach (var (type, operationId, tags) in operations)
         {
            pathItem.Operations.Add(type, new OpenApiOperation
            {
               OperationId = operationId,
               Summary = operationId,
               Tags = tags.Select(tag => Tag(tag)).ToList()
            });
         }

         return new KeyValuePair<string, OpenApiPathItem>(path, pathItem);
      }

      private static OpenApiDocument Document(IList<OpenApiTag> tags,
         IEnumerable<KeyValuePair<string, OpenApiPathItem>> paths)
      {
         var document = new OpenApiDocument
         {
            Info = new OpenApiInfo { Title = "Test", Version = "1.0" },
            Tags = tags,
            Paths = new OpenApiPaths()
         };

         foreach (var (path, item) in paths)
         {
            document.Paths.Add(path, item);
         }

         return document;
      }
   }
}
