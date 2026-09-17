# Changelog

Every released version of the tool, reconstructed from the history of the repository. The project
carries no tags, so a version covers the commits up to and including the one that raised
`semver.txt` to it. Only what a reader of the generated document, or a user of the command line,
can notice is listed here.

## 0.3.0

2026-09-17

### Added

- A last tab documenting the objects of the specification: every schema the document declares under
  a name and an operation reaches, one collapsible section each, described by the same table a
  request or a response body is described with. The `Object type` column of every tab links to the
  section documenting the object it names, and the tab links back to the first one. A schema the
  document declares and no operation reaches is not documented, and what the specification says
  about an object is written next to its name. A section collapses from the row naming the columns
  of its table, so collapsing every one of them leaves the list of the names.
- An object holding nothing, an enum or a simple type declared under a name, is described by one
  `<value>` row instead of by an empty table: its type, its format, its values and its range.
- The title of a body format names the object the body is and links to the section documenting it,
  `Body format: application/json (Pet)`. The table under it only names what the properties of that
  body hold.
- The first tab links to the last one, under the index of the operations.
- The index of the operations states the `operationId` of an operation next to its type, in one
  column, `PUT: updatePet`. An operation stating no `operationId` is stated by its type alone.
- The alternatives of a `oneOf` or an `anyOf` are documented as one row each, naming the
  alternative, and a schema written as several with `allOf` is documented as the one schema its
  parts describe. A `oneOf` used to be documented as nothing at all, and so was every composition
  of more than two schemas.
- `Object type` names the schemas a composition is written as, `All of (Cat, Dog)`, instead of
  calling it an object. The label is part of the translation, so the document says it in its own
  language.

### Fixed

- An array the document says nothing about, `type: array` with no `items`, crashed the tool instead
  of being documented as an array holding nothing the document describes.
- A response header stating a `$ref` was documented as a bare object. The reader of the
  specification leaves that one reference for the document to resolve, and now it is resolved: the
  header is documented as the type the components section declares.
- A worksheet whose name carries a character Excel refuses, which an `operationId` and a
  translation are both free to carry, made the tool throw. The character is replaced instead.
- A link to a worksheet whose name carries an apostrophe was rejected by Excel as an invalid
  reference.

## 0.2.1

2026-08-28

### Added

- The first tab groups the operations the way the specification groups them: by tag, and inside a
  tag by path, with the summary of every operation and a group that collapses from its header.
- The language of the generated document is chosen with `-l`, `--lang`. English and Polish ship
  with the tool, and any other language is a json file the option points at. The language decides
  the labels, the name of the info tab, `Yes` and `No`, and the culture the numbers and ranges
  taken from the specification are formatted with.
- An input file that breaks the OpenAPI standard is repaired instead of rejected, and every
  correction is printed. `--no-sanitize` restores the strict behaviour. Repaired: two paths with
  the same signature differing only in the name of their parameters, and a path parameter that
  does not match the path template.
- A `$ref` wrapped in a single element `allOf`, which is how a generator adds `nullable: true` to a
  reference, is documented with the type, name, format, enum values, default and example of the
  schema it references instead of as a bare `object`.
- `-d`, `--depth` caps how deep an object hierarchy is documented, for specifications with deeply
  nested or recursive schemas.

### Fixed

- A repeated `operationId` shorter than the worksheet name limit crashed the tool instead of
  producing a second worksheet.
- A fully qualified type name no longer prints itself over the Format column next to it.
- The description of a parameter is used when the parameter itself states one.

### Changed

- Runs on .NET 8.0 and 9.0. The end of life .NET 7.0 is gone.
- ClosedXML, Microsoft.OpenApi and Spectre.Console updated.

## 0.1.8

2024-11-03

### Added

- Default and object type columns in the schema tables.
- A default response is described by what the specification says about it.

## 0.1.7

2024-11-01

### Added

- `-v`, `--version` prints the version of the tool.

## 0.1.6

2024-10-27

### Added

- The operation information section names the `operationId` of the operation.

### Changed

- A worksheet is named after the `operationId` of its operation when the specification states one,
  and after the method and the path otherwise. A name already taken gets a numbered suffix.

## 0.1.5

2024-10-21

### Fixed

- A request or response body without a schema no longer crashes the schema table.

## 0.1.4

2024-10-21

### Fixed

- A request or response body without a schema no longer crashes the calculation of the tree depth.

## 0.1.3

2024-10-21

### Fixed

- Exception details containing square brackets no longer break the console output.

## 0.1.2

2024-10-20

### Added

- `-g`, `--debug` prints the whole exception instead of its message alone.
- Recursive data models are documented instead of ending in a stack overflow.

### Fixed

- Rows of a request and of a response group correctly.
- The description of an object type reads better.

## 0.1.1

2024-05-27

### Added

- An example column in the schema tables.
- The type of the items of an array is included in the request properties.

## 0.1.0

2024-05-22

### Added

- A property carries whether it is required.

### Fixed

- Response rows group correctly.

## 0.0.6

2024-05-17

### Added

- An enum column listing the values a property accepts.
- Rows of a request and of a response collapse into their section.
- Response headers get a section of their own.

### Changed

- An unexpected error prints its message instead of a stack trace.

## 0.0.5

2024-05-11

### Added

- Properties composed with `allOf` and `anyOf` are documented, and they count when the depth of the
  tree is calculated.
- Html tags are stripped from a description, and the text taken from the specification is trimmed.

### Fixed

- An operation without an `operationId` no longer breaks the naming of its worksheet.
- A file that cannot be written reports why instead of crashing.
- Property columns line up, and the description column is as wide as it needs to be.

## 0.0.4

2024-05-07

### Added

- `-n`, `--no-logo` runs the tool without the logo.
- A nullable column in the schema tables.

## 0.0.3

2024-04-18

### Changed

- The command line is built on Spectre.Console, which is where the help screen and the logo come
  from.

## 0.0.2

2024-04-13

### Added

- The input can be a URL the specification is downloaded from, not only a path to a file.

## 0.0.1

2024-04-07

### Added

- First release, published as a .NET global tool and as the `openapi2excel.core` library.
- An info tab describing the api and linking to every operation, and one tab per operation with its
  parameters, its request body and its responses, each schema documented as a tree of properties
  that collapses.
