using System.Text;
using ClosedXML.Excel;
using openapi2excel.core;
using openapi2excel.core.Lang;

namespace OpenApi2Excel.Tests
{
   /// <summary>
   /// The last worksheet, as it lands in the file: one collapsible section per object the operations
   /// use, each described by the table the request and the response bodies are described with, and
   /// the links between the Object type column of every worksheet and those sections.
   /// </summary>
   public class ObjectsWorksheetTest
   {
      private const string Document = """
         openapi: 3.0.1
         info:
           title: Devices
           version: v1
         paths:
           /devices:
             get:
               operationId: GetDevices
               parameters:
                 - name: kind
                   in: query
                   schema:
                     $ref: '#/components/schemas/DeviceKind'
               responses:
                 '200':
                   description: Ok
                   headers:
                     X-Rate-Limit:
                       schema:
                         $ref: '#/components/schemas/RateLimit'
                   content:
                     application/json:
                       schema:
                         type: array
                         items:
                           $ref: '#/components/schemas/Device'
           /devices/{id}:
             post:
               operationId: AddDevice
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
             Device:
               type: object
               required: [ name ]
               properties:
                 name:
                   type: string
                   description: Name of the device
                 kind:
                   $ref: '#/components/schemas/DeviceKind'
                 location:
                   $ref: '#/components/schemas/Location'
             Location:
               type: object
               properties:
                 latitude:
                   type: number
                   format: double
                 longitude:
                   type: number
                   format: double
             DeviceKind:
               type: string
               description: What the device is
               enum: [ sensor, gateway ]
             RateLimit:
               type: integer
               format: int32
             NeverUsed:
               type: object
               properties:
                 nothing:
                   type: string
         """;

      private const string Shapes = """
         openapi: 3.0.1
         info:
           title: Shapes
           version: v1
         paths:
           /pets:
             get:
               operationId: GetPets
               responses:
                 '200':
                   description: Ok
                   content:
                     application/json:
                       schema:
                         $ref: '#/components/schemas/PetList'
             post:
               operationId: AddPet
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
           /remote:
             get:
               operationId: GetRemote
               responses:
                 '200':
                   description: Ok
                   content:
                     application/json:
                       schema:
                         type: object
                         properties:
                           here:
                             $ref: '#/components/schemas/Error'
                           there:
                             $ref: 'common.yaml#/components/schemas/Error'
         components:
           schemas:
             PetList:
               type: array
               items:
                 $ref: '#/components/schemas/Cat'
             AnyPet:
               oneOf:
                 - $ref: '#/components/schemas/Cat'
                 - $ref: '#/components/schemas/Dog'
             Merged:
               allOf:
                 - $ref: '#/components/schemas/Cat'
                 - $ref: '#/components/schemas/Dog'
             Cat:
               type: object
               properties:
                 meows:
                   type: boolean
             Dog:
               type: object
               properties:
                 barks:
                   type: boolean
             Error:
               type: object
               properties:
                 code:
                   type: integer
         """;

      [Fact]
      public void Objects_worksheet_comes_last_and_has_a_section_for_every_object_an_operation_uses()
      {
         using var workbook = Generate("en");

         Assert.Equal(["Info", "GetDevices", "AddDevice", "Objects"],
            workbook.Worksheets.Select(worksheet => worksheet.Name));

         // In the order a reader looks a name up, and only the objects the operations reach:
         // NeverUsed is declared by the document and used by nothing.
         Assert.Equal(["Device", "DeviceKind", "Location", "RateLimit"], SectionTitles(workbook.Worksheet("Objects")));
      }

      [Fact]
      public void Object_is_described_by_the_table_that_describes_a_request_body()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");

         // The name of the object, the header of the table, then one row per property, with a
         // nested object opened under the property holding it.
         Assert.Equal("Device", objects.Cell("A4").GetString());
         Assert.Equal(["Name", "", "", "", "Type", "Object type", "Format", "Length", "Required",
               "Nullable", "Range", "Pattern", "Enum", "Deprecated", "Default", "Example", "Description"],
            Enumerable.Range(1, 17).Select(column => objects.Row(5).Cell(column).GetString()));

         Assert.Equal(["name", "string", "Yes", "Name of the device"],
            Row(objects, 6, [1, 5, 9, 17]));
         Assert.Equal(["kind", "string", "DeviceKind", "sensor, gateway"],
            Row(objects, 7, [1, 5, 6, 13]));
         Assert.Equal(["location", "object", "Location"], Row(objects, 8, [1, 5, 6]));
         Assert.Equal(["", "latitude", "number", "double"], Row(objects, 9, [1, 2, 5, 7]));
         Assert.Equal(["", "longitude", "number", "double"], Row(objects, 10, [1, 2, 5, 7]));
      }

      [Fact]
      public void Every_section_collapses_from_the_row_naming_the_columns_of_its_table()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");

         Assert.Equal(XLOutlineSummaryVLocation.Top, objects.Outline.SummaryVLocation);

         // The link and the header of the worksheet, then Device and the table describing it. The
         // collapsing starts at the row naming the columns, one row below the name of the object.
         Assert.Equal([0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 0, 0, 1],
            Enumerable.Range(1, 13).Select(row => objects.Row(row).OutlineLevel));
      }

      [Fact]
      public void Name_of_an_object_is_visible_whatever_a_reader_collapses()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");

         // A reader looks an object up by its name, so no group holds a name and collapsing every
         // section leaves the list of them. Everything below a name collapses into it.
         foreach (var titleRow in SectionRows(objects))
         {
            Assert.Equal(0, objects.Row(titleRow).OutlineLevel);
            Assert.Equal("Name", objects.Cell(titleRow + 1, 1).GetString());
            Assert.Equal(1, objects.Row(titleRow + 1).OutlineLevel);
         }

         Assert.Equal(["Device", "DeviceKind", "Location", "RateLimit"], SectionTitles(objects));
      }

      [Fact]
      public void Title_of_a_section_is_told_apart_from_the_table_below_it()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");

         // The title is bold and filled the whole width of the table, the way a body format is.
         Assert.True(objects.Cell("A4").Style.Font.Bold);
         Assert.Equal(XLColor.LightGray, objects.Cell("A4").Style.Fill.BackgroundColor);
         Assert.Equal(XLColor.LightGray, objects.Cell("Q4").Style.Fill.BackgroundColor);

         // The header of the table carries the line under it, the properties carry nothing.
         Assert.True(objects.Cell("A5").Style.Font.Bold);
         Assert.Equal(XLBorderStyleValues.Medium, objects.Cell("Q5").Style.Border.BottomBorder);
         Assert.False(objects.Cell("A6").Style.Font.Bold);
      }

      [Fact]
      public void Objects_worksheet_links_back_to_the_first_one()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");

         Assert.Equal("<<<<<", objects.Cell("A1").GetString());
         Assert.Equal("'Info'!A1", objects.Cell("A1").GetHyperlink().InternalAddress);
         Assert.Equal("OBJECTS", objects.Cell("A3").GetString());
      }

      [Fact]
      public void First_worksheet_links_to_the_objects_worksheet()
      {
         using var workbook = Generate("en");
         var info = workbook.Worksheet("Info");
         var lastRow = info.LastRowUsed()!.RowNumber();

         // The objects worksheet links back to the first one, and nothing else on the first one
         // would tell a reader that it exists at all.
         Assert.Equal("OBJECTS", info.Cell(lastRow, 1).GetString());
         Assert.Equal("'Objects'!A1", info.Cell(lastRow, 1).GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Object_type_of_an_operation_links_to_the_section_documenting_that_object()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");
         var getDevices = workbook.Worksheet("GetDevices");

         // A parameter, a response header, the body of the response and a property of that body all
         // name a schema, and each of them walks to the section describing it.
         Assert.Equal($"'Objects'!A{SectionRow(objects, "DeviceKind")}",
            CellNaming(getDevices, "DeviceKind").GetHyperlink().InternalAddress);
         Assert.Equal($"'Objects'!A{SectionRow(objects, "RateLimit")}",
            CellNaming(getDevices, "RateLimit").GetHyperlink().InternalAddress);
         Assert.Equal($"'Objects'!A{SectionRow(objects, "Device")}",
            CellNaming(getDevices, "Device").GetHyperlink().InternalAddress);
         Assert.Equal($"'Objects'!A{SectionRow(objects, "Location")}",
            CellNaming(workbook.Worksheet("AddDevice"), "Location").GetHyperlink().InternalAddress);

         // Every cell naming an object links, none of them is left as plain text.
         Assert.All(CellsNaming(getDevices, "Device", "DeviceKind", "Location", "RateLimit"),
            cell => Assert.True(cell.HasHyperlink, $"{cell.Address} is not a link"));
      }

      [Fact]
      public void Body_of_an_operation_names_the_object_it_is_and_links_to_it()
      {
         using var workbook = Generate("en");
         var addDevice = workbook.Worksheet("AddDevice");
         var objects = workbook.Worksheet("Objects");

         // The table under a body format describes the properties of the body, and nothing else
         // says what the body itself is.
         var title = addDevice.CellsUsed(cell => cell.GetString().StartsWith("Body format:")).Single();
         Assert.Equal("Body format: application/json (Device)", title.GetString());
         Assert.Equal($"'Objects'!A{SectionRow(objects, "Device")}", title.GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Body_holding_a_list_of_an_object_is_not_named_after_that_object()
      {
         using var workbook = Generate("en");
         var getDevices = workbook.Worksheet("GetDevices");

         // The response is a list of devices, and a list of devices is not a device. The row of the
         // array, right under the header of the table, is what names the object it holds.
         var title = getDevices.CellsUsed(cell => cell.GetString().StartsWith("Body format:")).Single();
         Assert.Equal("Body format: application/json", title.GetString());
         Assert.False(title.HasHyperlink);

         var arrayRow = getDevices.Row(title.Address.RowNumber + 2);
         var named = arrayRow.CellsUsed(cell => cell.HasHyperlink).Single();
         Assert.Equal("<array>", arrayRow.Cell(1).GetString());
         Assert.Equal("Device", named.GetString());
         Assert.Equal($"'Objects'!A{SectionRow(workbook.Worksheet("Objects"), "Device")}",
            named.GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Section_says_what_the_specification_says_about_the_object()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");
         var row = SectionRow(objects, "DeviceKind");

         // Next to the name, in the column every description of the table below already uses.
         Assert.Equal("What the device is", objects.Cell(row, 17).GetString());
         Assert.Equal("", objects.Cell(SectionRow(objects, "Location"), 17).GetString());
      }

      [Fact]
      public void Object_type_inside_a_section_links_to_the_section_of_the_object_it_names()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");

         // Device holds a Location, and its row walks to the section describing one.
         Assert.Equal($"'Objects'!A{SectionRow(objects, "Location")}",
            objects.Cell("F8").GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Link_is_painted_the_way_a_link_is_painted()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");

         // Setting a hyperlink changes nothing about how a cell looks, and a reader has to see
         // there is one to follow.
         Assert.Equal(XLColor.FromArgb(0x05, 0x63, 0xC1), objects.Cell("F8").Style.Font.FontColor);
         Assert.Equal(XLFontUnderlineValues.Single, objects.Cell("F8").Style.Font.Underline);

         // A cell naming no object is left alone.
         Assert.NotEqual(XLFontUnderlineValues.Single, objects.Cell("F6").Style.Font.Underline);
      }

      [Fact]
      public void Object_with_no_property_of_its_own_still_describes_itself()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");
         var row = SectionRow(objects, "DeviceKind") + 2;

         // An enum has no property to list, and the one row of its table is what it is worth
         // reading: the type and the values of the schema itself. What the specification says about
         // it is written once, next to its name, and not again here.
         Assert.Equal(["<value>", "string", "sensor, gateway", ""], Row(objects, row, [1, 5, 13, 17]));

         // The name of the object is the title of the section the row already sits in, so the
         // Object type column does not repeat it and links nowhere.
         Assert.Equal("", objects.Cell(row, 6).GetString());
         Assert.False(objects.Cell(row, 6).HasHyperlink);
      }

      [Fact]
      public void Reference_the_reader_leaves_unresolved_is_documented_from_the_components_section()
      {
         using var workbook = Generate("en");
         var objects = workbook.Worksheet("Objects");
         var row = SectionRow(objects, "RateLimit") + 2;

         // A $ref under a response header reaches the document as the name of a schema and nothing
         // else. What the components section declares is the schema itself.
         Assert.Equal(["<value>", "integer", "int32"], Row(objects, row, [1, 5, 7]));

         // The worksheet of the operation says the same. Two worksheets of one document cannot
         // disagree about what a type is, and the link between them is what makes it checkable.
         var getDevices = workbook.Worksheet("GetDevices");
         var header = getDevices.CellsUsed(cell => cell.GetString() == "X-Rate-Limit").Single();
         Assert.Equal(["X-Rate-Limit", "integer", "RateLimit", "int32"],
            Row(getDevices, header.Address.RowNumber, [1, 6, 7, 8]));
      }

      [Fact]
      public void Name_of_the_worksheet_and_its_header_follow_the_language()
      {
         using var workbook = Generate("pl");
         var objects = workbook.Worksheet("Obiekty");

         Assert.Equal("OBIEKTY", objects.Cell("A3").GetString());
         Assert.Equal("'Informacje'!A1", objects.Cell("A1").GetHyperlink().InternalAddress);

         // The names of the specification are not translated, and the links still find them.
         Assert.Equal(["Device", "DeviceKind", "Location", "RateLimit"], SectionTitles(objects));
         Assert.Equal($"'Obiekty'!A{SectionRow(objects, "Location")}",
            objects.Cell("F8").GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Specification_naming_no_object_at_all_has_no_objects_worksheet()
      {
         const string inlineOnly = """
            openapi: 3.0.1
            info:
              title: Nothing named
              version: v1
            paths:
              /health:
                get:
                  operationId: GetHealth
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            type: object
                            properties:
                              ok:
                                type: boolean
            """;

         using var workbook = Generate("en", inlineOnly);

         Assert.Equal(["Info", "GetHealth"], workbook.Worksheets.Select(worksheet => worksheet.Name));
      }

      [Fact]
      public void Array_the_document_says_nothing_about_names_no_object_and_does_not_throw()
      {
         const string arrayWithoutItems = """
            openapi: 3.0.1
            info:
              title: Bare array
              version: v1
            paths:
              /things:
                get:
                  operationId: GetThings
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            type: object
                            properties:
                              values:
                                type: array
            """;

         // Regression: an array with no items used to throw while the Object type column was
         // written, which is the column naming what the array holds.
         using var workbook = Generate("en", arrayWithoutItems);
         var things = workbook.Worksheet("GetThings");
         var row = things.CellsUsed(cell => cell.GetString() == "values").Single().Address.RowNumber;

         Assert.Equal(["values", "array", ""], Row(things, row, [1, 4, 5]));
         Assert.Equal(["Info", "GetThings"], workbook.Worksheets.Select(worksheet => worksheet.Name));
      }

      [Fact]
      public void Operation_named_the_way_the_objects_worksheet_is_named_does_not_take_its_name()
      {
         const string collision = """
            openapi: 3.0.1
            info:
              title: Collision
              version: v1
            paths:
              /nodes:
                get:
                  operationId: Objects
                  responses:
                    '200':
                      description: Ok
                      content:
                        application/json:
                          schema:
                            type: array
                            items:
                              $ref: '#/components/schemas/Node'
            components:
              schemas:
                Node:
                  type: object
                  properties:
                    name:
                      type: string
            """;

         using var workbook = Generate("en", collision);

         // Excel refuses two worksheets of the same name, and the links follow the name the
         // worksheet ends up with.
         Assert.Equal(["Info", "Objects", "Objects_2"], workbook.Worksheets.Select(worksheet => worksheet.Name));
         Assert.Equal("'Objects_2'!A4",
            CellNaming(workbook.Worksheet("Objects"), "Node").GetHyperlink().InternalAddress);

         var info = workbook.Worksheet("Info");
         Assert.Equal("'Objects_2'!A1",
            info.Cell(info.LastRowUsed()!.RowNumber(), 1).GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Object_holding_itself_is_documented_once_and_walks_back_to_its_own_section()
      {
         const string recursive = """
            openapi: 3.0.1
            info:
              title: Recursive
              version: v1
            paths:
              /nodes:
                get:
                  operationId: GetNodes
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
                    name:
                      type: string
                    parent:
                      $ref: '#/components/schemas/Node'
            """;

         using var workbook = Generate("en", recursive, maxDepth: 3);
         var objects = workbook.Worksheet("Objects");

         // One section, however many times the schema holds itself, and the tree of that section
         // stops at the depth the options allow instead of running forever.
         Assert.Equal(["Node"], SectionTitles(objects));
         Assert.Equal(["name", "parent"], new[] { 6, 7 }.Select(row => objects.Cell(row, 1).GetString()));
         Assert.Equal(["name", "parent"], new[] { 8, 9 }.Select(row => objects.Cell(row, 2).GetString()));
         Assert.Equal(9, objects.LastRowUsed()!.RowNumber());

         // A property naming the object of the section it sits in walks back to the top of it.
         Assert.Equal("'Objects'!A4", objects.Cell("F7").GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Alternatives_of_a_schema_are_documented_as_rows_naming_them()
      {
         using var workbook = Generate("en", Shapes);
         var objects = workbook.Worksheet("Objects");
         var anyPet = SectionRow(objects, "AnyPet");

         // oneOf and anyOf are alternatives, not parts. Writing their properties together would
         // read as one object holding all of them.
         Assert.Equal(["Cat", "", "Dog", ""],
            Enumerable.Range(anyPet + 2, 4).Select(row => objects.Cell(row, 1).GetString()));
         Assert.Equal($"'Objects'!A{SectionRow(objects, "Cat")}",
            objects.Cell(anyPet + 2, 6).GetHyperlink().InternalAddress);
         Assert.Equal(["meows", "barks"],
            new[] { anyPet + 3, anyPet + 5 }.Select(row => objects.Cell(row, 2).GetString()));
      }

      [Fact]
      public void Schema_written_in_parts_is_documented_as_the_one_schema_it_describes()
      {
         using var workbook = Generate("en", Shapes);
         var objects = workbook.Worksheet("Objects");
         var merged = SectionRow(objects, "Merged");

         // allOf is one schema written in several, and every part describes the same object.
         Assert.Equal(["meows", "barks"],
            new[] { merged + 2, merged + 3 }.Select(row => objects.Cell(row, 1).GetString()));
      }

      [Fact]
      public void List_the_document_declares_under_a_name_is_a_section_of_its_own()
      {
         using var workbook = Generate("en", Shapes);
         var objects = workbook.Worksheet("Objects");
         var petList = SectionRow(objects, "PetList");

         // The section documents the list, and the row of the array names what the list holds.
         Assert.Equal("<array>", objects.Cell(petList + 2, 1).GetString());
         Assert.Equal($"'Objects'!A{SectionRow(objects, "Cat")}",
            objects.Cell(petList + 2, 6).GetHyperlink().InternalAddress);

         // The body of the operation is that list, and says so.
         var title = workbook.Worksheet("GetPets")
            .CellsUsed(cell => cell.GetString().StartsWith("Body format:")).Single();
         Assert.Equal("Body format: application/json (PetList)", title.GetString());
         Assert.Equal($"'Objects'!A{petList}", title.GetHyperlink().InternalAddress);
      }

      [Fact]
      public void Reference_to_another_document_names_a_schema_this_one_cannot_describe()
      {
         using var workbook = Generate("en", Shapes);
         var objects = workbook.Worksheet("Objects");
         var remote = workbook.Worksheet("GetRemote");

         // Both properties name an Error, and only one of them is the Error this document declares.
         // Linking the other one to it would tell the reader the remote schema holds a code.
         Assert.Single(SectionTitles(objects), title => title == "Error");
         Assert.Equal(["here", "Error"], Row(remote, 13, [1, 6]));
         Assert.True(remote.Cell(13, 6).HasHyperlink);
         Assert.Equal(["there", "Error"], Row(remote, 15, [1, 6]));
         Assert.False(remote.Cell(15, 6).HasHyperlink);
      }

      [Fact]
      public void Title_of_a_body_is_clickable_across_the_text_a_reader_sees()
      {
         using var workbook = Generate("en");
         var addDevice = workbook.Worksheet("AddDevice");
         var title = addDevice.CellsUsed(cell => cell.GetString().StartsWith("Body format:")).Single();
         var row = title.Address.RowNumber;
         var address = title.GetHyperlink().InternalAddress;

         // The title spills over the cells to its right, which hold nothing, and the cell holding
         // the text is two characters wide. All of them walk to the section.
         Assert.All(Enumerable.Range(2, 3),
            column => Assert.Equal(address, addDevice.Cell(row, column).GetHyperlink().InternalAddress));
      }

      [Fact]
      public void Depth_the_options_stop_at_does_not_turn_an_object_into_a_value()
      {
         using var workbook = Generate("en", Document, maxDepth: 1);
         var objects = workbook.Worksheet("Objects");
         var device = SectionRow(objects, "Device");

         // The properties of Device are cut off by the depth, and a row calling it a value would
         // say the opposite of what the specification says.
         Assert.Equal("", objects.Cell(device + 2, 1).GetString());

         // A schema that really holds nothing still describes itself. The table is narrower here,
         // the depth decides how many columns the names of a tree need.
         var deviceKind = SectionRow(objects, "DeviceKind");
         Assert.Equal("<value>", objects.Cell(deviceKind + 2, 1).GetString());
         Assert.Contains("string", objects.Row(deviceKind + 2).CellsUsed().Select(cell => cell.GetString()));
      }

      [Fact]
      public void Name_a_translation_gives_the_worksheet_is_made_into_one_excel_accepts()
      {
         using var workbook = Generate(new Translation { ObjectsSheetName = "L'Objets [all]" });
         var objects = workbook.Worksheets.Last();
         var info = workbook.Worksheet("Info");

         // Excel refuses a bracket in the name of a worksheet, and quotes that name in every link
         // to it, where an apostrophe ends the name unless it is doubled.
         Assert.Equal("L'Objets -all-", objects.Name);
         Assert.Equal("'L''Objets -all-'!A1",
            info.Cell(info.LastRowUsed()!.RowNumber(), 1).GetHyperlink().InternalAddress);
         Assert.Equal($"'L''Objets -all-'!A{SectionRow(objects, "Device")}",
            CellNaming(workbook.Worksheet("GetDevices"), "Device").GetHyperlink().InternalAddress);
      }

      /// <summary>Titles of the sections, in the order the worksheet writes them.</summary>
      private static List<string> SectionTitles(IXLWorksheet worksheet)
         => SectionRows(worksheet).Select(row => worksheet.Cell(row, 1).GetString()).ToList();

      private static int SectionRow(IXLWorksheet worksheet, string objectName)
         => SectionRows(worksheet).First(row => worksheet.Cell(row, 1).GetString() == objectName);

      /// <summary>
      /// Rows carrying the title of a section: nothing groups the sections together, so a title is
      /// outside every group and the table describing it is one level in. The empty row separating
      /// two sections is outside every group as well, and carries no title.
      /// </summary>
      private static IEnumerable<int> SectionRows(IXLWorksheet worksheet)
         => Enumerable.Range(4, (worksheet.LastRowUsed()?.RowNumber() ?? 3) - 3)
            .Where(row => worksheet.Row(row).OutlineLevel == 0
                          && !string.IsNullOrEmpty(worksheet.Cell(row, 1).GetString()));

      private static string[] Row(IXLWorksheet worksheet, int rowNumber, int[] columns)
         => columns.Select(column => worksheet.Cell(rowNumber, column).GetString()).ToArray();

      /// <summary>The one cell of a worksheet naming an object, wherever the table puts it.</summary>
      private static IXLCell CellNaming(IXLWorksheet worksheet, string objectName)
         => CellsNaming(worksheet, objectName).First();

      /// <summary>
      /// The cells naming one of the objects, the values of the Object type column. A header is
      /// bold and is left out: the parameters of an operation are described by a column called
      /// Location, which is also the name of an object of the document.
      /// </summary>
      private static List<IXLCell> CellsNaming(IXLWorksheet worksheet, params string[] objectNames)
         => worksheet.CellsUsed(cell => objectNames.Contains(cell.GetString())
                                        && cell.Address.ColumnNumber > 1
                                        && !cell.Style.Font.Bold)
            .OrderBy(cell => cell.Address.RowNumber)
            .ToList();

      private static XLWorkbook Generate(Translation translation, string document = Document)
         => Generate(new OpenApiDocumentationOptions { Translation = translation }, document);

      private static XLWorkbook Generate(string language, string document = Document, int maxDepth = 10)
         => Generate(
            new OpenApiDocumentationOptions
            {
               Translation = TranslationsSelector.Load(language).Translation,
               MaxDepth = maxDepth
            },
            document);

      private static XLWorkbook Generate(OpenApiDocumentationOptions options, string document)
      {
         var outputFile = Path.Combine(Path.GetTempPath(), $"openapi2excel-tests-{Guid.NewGuid():N}.xlsx");
         try
         {
            using var input = new MemoryStream(Encoding.UTF8.GetBytes(document));
            OpenApiDocumentationGenerator.GenerateDocumentation(input, outputFile, options)
               .GetAwaiter().GetResult();

            using var generated = File.OpenRead(outputFile);
            return new XLWorkbook(generated);
         }
         finally
         {
            File.Delete(outputFile);
         }
      }
   }
}
