using Microsoft.OpenApi.Models;
using Microsoft.OpenApi.Readers;
using openapi2excel.core.Builders;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// Gate for what lands on the objects worksheet. Which schemas are objects, and in which order,
   /// is decided here and only rendered by the worksheet builder, so it is asserted here instead of
   /// through cell addresses.
   /// </summary>
   public class ObjectsCatalogTest
   {
      [Fact]
      public void Every_place_an_operation_names_a_schema_is_an_object()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Named everywhere, version: v1 }
            paths:
              /devices:
                get:
                  parameters:
                    - name: kind
                      in: query
                      schema:
                        $ref: '#/components/schemas/InAParameter'
                  responses:
                    '200':
                      description: Ok
                      headers:
                        X-Rate-Limit:
                          schema:
                            $ref: '#/components/schemas/InAResponseHeader'
                      content:
                        application/json:
                          schema:
                            $ref: '#/components/schemas/InAResponseBody'
                post:
                  requestBody:
                    content:
                      application/json:
                        schema:
                          $ref: '#/components/schemas/InARequestBody'
                  responses:
                    '200':
                      description: Ok
            components:
              schemas:
                InAParameter: { type: string }
                InAResponseHeader: { type: integer }
                InAResponseBody: { type: object, properties: { a: { type: string } } }
                InARequestBody: { type: object, properties: { b: { type: string } } }
            """);

         Assert.Equal(["InAParameter", "InARequestBody", "InAResponseBody", "InAResponseHeader"], Names(document));
      }

      [Fact]
      public void Schema_named_only_by_another_schema_is_an_object_as_well()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Nested, version: v1 }
            paths:
              /devices:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            $ref: '#/components/schemas/Device'
            components:
              schemas:
                Device:
                  type: object
                  properties:
                    location:
                      $ref: '#/components/schemas/Location'
                Location:
                  type: object
                  properties:
                    country:
                      $ref: '#/components/schemas/Country'
                Country: { type: string, enum: [ pl, en ] }
            """);

         // The section describing Device names Location, and the one describing Location names
         // Country, so both of them need a section of their own to link to.
         Assert.Equal(["Country", "Device", "Location"], Names(document));
      }

      [Fact]
      public void Schema_the_document_declares_and_no_operation_reaches_is_not_an_object()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Unused, version: v1 }
            paths:
              /devices:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            $ref: '#/components/schemas/Used'
            components:
              schemas:
                Used: { type: object, properties: { a: { type: string } } }
                NeverUsed: { type: object, properties: { b: { type: string } } }
            """);

         Assert.Equal(["Used"], Names(document));
      }

      [Fact]
      public void Object_used_by_several_operations_is_one_object()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Shared, version: v1 }
            paths:
              /devices:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            $ref: '#/components/schemas/Device'
                post:
                  requestBody:
                    content:
                      application/json:
                        schema:
                          $ref: '#/components/schemas/Device'
                  responses:
                    '200':
                      description: Ok
            components:
              schemas:
                Device: { type: object, properties: { a: { type: string } } }
            """);

         Assert.Equal(["Device"], Names(document));
      }

      [Fact]
      public void Array_and_a_reference_wrapped_to_carry_a_keyword_name_the_schema_they_hold()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Wrapped, version: v1 }
            paths:
              /devices:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            type: array
                            items:
                              $ref: '#/components/schemas/Device'
            components:
              schemas:
                Device:
                  type: object
                  properties:
                    kind:
                      allOf: [ { $ref: '#/components/schemas/DeviceKind' } ]
                      nullable: true
                    tags:
                      type: array
                      items:
                        $ref: '#/components/schemas/Tag'
                DeviceKind: { type: string, enum: [ sensor ] }
                Tag: { type: object, properties: { name: { type: string } } }
            """);

         // The Object type column names what an array holds and what a wrapper wraps, so those are
         // the names the sections have to carry.
         Assert.Equal(["Device", "DeviceKind", "Tag"], Names(document));
      }

      [Fact]
      public void Object_declared_in_place_names_nothing_and_is_no_object_of_its_own()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Inline, version: v1 }
            paths:
              /devices:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            type: object
                            properties:
                              location:
                                type: object
                                properties:
                                  latitude: { type: number }
            """);

         Assert.Empty(ObjectsCatalog.Collect(document));
      }

      [Fact]
      public void Array_the_document_says_nothing_about_holds_no_object()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Bare array, version: v1 }
            paths:
              /things:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            type: object
                            properties:
                              values: { type: array }
                              thing: { $ref: '#/components/schemas/Thing' }
            components:
              schemas:
                Thing: { type: array }
            """);

         // An array is named by what it holds, and these hold nothing the document describes, so
         // the Object type column names nothing to link to. The one the document declares under a
         // name is still a schema of the specification, and the worksheet documents it as one.
         Assert.Equal(["Thing"], Names(document));
      }

      [Fact]
      public void Schema_holding_itself_is_collected_once()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Recursive, version: v1 }
            paths:
              /nodes:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            $ref: '#/components/schemas/Node'
            components:
              schemas:
                Node:
                  type: object
                  properties:
                    parent:
                      $ref: '#/components/schemas/Node'
                    children:
                      type: array
                      items:
                        $ref: '#/components/schemas/Node'
            """);

         Assert.Equal(["Node"], Names(document));
      }

      [Fact]
      public void Reference_to_another_document_names_a_schema_this_one_cannot_describe()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: External, version: v1 }
            paths:
              /errors:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            type: object
                            properties:
                              here: { $ref: '#/components/schemas/Error' }
                              there: { $ref: 'common.yaml#/components/schemas/Error' }
            components:
              schemas:
                Error: { type: object, properties: { code: { type: integer } } }
            """);

         // The document never read common.yaml. Documenting the local Error as the remote one, only
         // because they carry the same name, would invent everything the section says about it.
         var error = ObjectsCatalog.Collect(document).Single();
         Assert.Equal("Error", error.Name);
         Assert.Equal(["code"], error.Schema.Properties.Keys);
      }

      [Fact]
      public void List_the_document_declares_under_a_name_is_an_object_of_its_own()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Lists, version: v1 }
            paths:
              /pets:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            $ref: '#/components/schemas/PetList'
            components:
              schemas:
                PetList: { type: array, items: { $ref: '#/components/schemas/Pet' } }
                Pet: { type: object, properties: { name: { type: string } } }
            """);

         // A reader of the specification looks a list up by the name it is declared with, and the
         // Object type column of a row holding one names what the list holds.
         Assert.Equal(["Pet", "PetList"], Names(document));
      }

      [Fact]
      public void Alternatives_and_parts_of_a_schema_are_objects_as_well()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Compositions, version: v1 }
            paths:
              /pets:
                post:
                  requestBody:
                    content:
                      application/json:
                        schema:
                          $ref: '#/components/schemas/AnyPet'
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            $ref: '#/components/schemas/Merged'
            components:
              schemas:
                AnyPet:
                  oneOf:
                    - $ref: '#/components/schemas/Cat'
                    - $ref: '#/components/schemas/Dog'
                Merged:
                  allOf:
                    - $ref: '#/components/schemas/Cat'
                    - $ref: '#/components/schemas/Dog'
                Cat: { type: object, properties: { meows: { type: boolean } } }
                Dog: { type: object, properties: { barks: { type: boolean } } }
            """);

         Assert.Equal(["AnyPet", "Cat", "Dog", "Merged"], Names(document));
      }

      [Fact]
      public void Objects_come_in_the_order_a_reader_looks_a_name_up()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Sorted, version: v1 }
            paths:
              /devices:
                get:
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            $ref: '#/components/schemas/zeta'
            components:
              schemas:
                zeta:
                  type: object
                  properties:
                    a: { $ref: '#/components/schemas/Alpha' }
                    b: { $ref: '#/components/schemas/beta' }
                    c: { $ref: '#/components/schemas/Gamma' }
                Alpha: { type: object, properties: { x: { type: string } } }
                beta: { type: object, properties: { x: { type: string } } }
                Gamma: { type: object, properties: { x: { type: string } } }
            """);

         // Sorted by name and not by case, the way a list is read, and not in the order the
         // document happens to declare them.
         Assert.Equal(["Alpha", "beta", "Gamma", "zeta"], Names(document));
      }

      [Fact]
      public void Schema_the_components_section_declares_is_the_one_documented()
      {
         var document = Read("""
            openapi: 3.0.1
            info: { title: Header, version: v1 }
            paths:
              /devices:
                get:
                  responses:
                    '200':
                      description: Ok
                      headers:
                        X-Rate-Limit:
                          schema:
                            $ref: '#/components/schemas/RateLimit'
            components:
              schemas:
                RateLimit: { type: integer, format: int32 }
            """);

         // A $ref under a response header reaches the document as a name and nothing else, while
         // the components section always declares the schema itself.
         var rateLimit = ObjectsCatalog.Collect(document).Single();
         Assert.Equal("RateLimit", rateLimit.Name);
         Assert.Equal("integer", rateLimit.Schema.Type);
         Assert.Equal("int32", rateLimit.Schema.Format);
      }

      private static IEnumerable<string> Names(OpenApiDocument document)
         => ObjectsCatalog.Collect(document).Select(documented => documented.Name);

      private static OpenApiDocument Read(string document)
         => new OpenApiStringReader().Read(document, out _);
   }
}
