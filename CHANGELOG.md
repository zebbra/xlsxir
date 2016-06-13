# Change Log

## 1.3.0

- Added ability to parse multiple worksheets via `Xlsxir.multi_extract/3` which returns a unique table identifier for each ETS process created, enabling the user to access parsed data from multiple worksheets simultaneously. 
- Created an `Xlsxir.TableId` module which controls an agent process that temporarily holds a table identifier during the extraction process.
- Refactored `Xlsxir` access functions to work with `Xlsxir.multi_extract/3` whereby a table identifier is passed through the various functions to specify which ETS process is to be accessed. 
- Refactored `Xlsxir.SaxParser`, `Xlsxir.ParseWorksheet` and `Xlsxir.Worksheet` modules to support new functionality.
- Updated documentation and tests
- Fixed a few minor bugs that were generating warning messages. 

## 1.2.1

- Removed Ex-Doc and Earmark dependencies from Hex.
- Added Change Log link to Hex.
- Minor doc changes and bug fixes.

## 1.2.0

- Added `Xlsxir` access function `Xlsxir.get_mda/0` that accesses `:worksheet` ETS table and returns an indexed map which functions like a multi-dimensional array in other languages.

## 1.1.0

- Modified the way rows are saved to the `:worksheet` ETS table. Replaced the generic index with the actual row number to allow for performance imporovement of supporting  `Xlsxir` access functions.
- Refactored `Xlsxir` access functions to improve performance.
- Created `Xlsxir.get_info/1` function which returns number of rows, columns and cells. 
- Various minor modifications to docs. 

## 1.0.0

Major changes in version 1.0.0 (non-backwards compatible) to improve performance and incorporate new functionality, including: 

- Refactored the `Xlsxir.Unzip` module to extract `.xlsx` contents to file instead of memory to improve memory usage. The following functions were created to support this functionality:
    * `Xlsxir.Unzip.extract_xml_to_file/2` - Extracts necessary files to a `./temp` directory for use during the parsing process
    * `Xlsxir.Unzip.delete_dir/1` - Deletes './temp' directory and all of its contents
- Implemented Simple API for XML (SAX) parsing functionalty via the [Erlsom](https://github.com/willemdj/erlsom) Erlang library to improve performance and allow support for large `.xlsx` files. The `SweetXml` parsing library has been deprecated from `Xlsxir` and is no longer utilized in v1.0.0.    
- Implemented Erlang Term Storage (ETS) for temporary storage of extracted data.
- Replaced `option` argument from the initial extract function (`Xlsxir.extract/3`) with `timer` which is a boolean flag that controls `Xlsxir.Timer` functionality. Data is no longer returned via `Xlsxir.extract/3` and is instead stored in an ETS process.
- Implemented various functions for accessing the extracted data:
    * `Xlsxir.get_list/0` - Return entire worksheet data in the form of a list of row lists 
    * `Xlsxir.get_map/0` - Return entire worksheet data in the form of a map of cell names and values
    * `Xlsxir.get_cell/1` - Return value of specified cell
    * `Xlsxir.get_row/1` - Return values of specified row
    * `Xlsxir.get_col/1` - Return values of specified column
- Implemented `Xlsxir.close/0` function to allow the deletion of the ETS process containing extracted worksheet data to free up memory.
- Implemented `Xlsxir.Timer` module for tracking elapsed time of extraction process.
- Changed cell references from `atoms` to `strings` due to Elixir `atom` limitations (i.e. `:A1` to `"A1"`).
- Updated documentation and testing to incorporate changes.

## 0.0.5

- Minor bug fixes and documentation updates.

## 0.0.4

- Expanded coverage of Office Open XML standard `numFmt` (Standard Number Format). The `formatCode` for a standard `numFmt` is implied rather than explicitly identified in the XML file.
- Implemented support for Office Open XML custom `numFmt` (Custom Number Format) utilizing the `formatCode` explicitly identified in the XML file. 
- Added `Number Styles` documentation covering standard and custom `numFmt` and how to manually add an unsupported `numFmt`.
- Fixed issue resulting when no strings exist in a worksheet and therefore there is no `sharedStrings.xml` file (`:file_not_found` error).

## 0.0.3

- Fixed issue related to strings that contain special characters. Refactored `Xlsxir.Parse.shared_strings/1` to properly parse strings with special characters.

## 0.0.2

- Refactored `Xlsxir.Parse` functions to improve `extract` performance on larger files.
- Expanded documentation and test coverage.
- Completed `Xlsxir.ConvertDate` module.

## 0.0.1

- Initial draft. Functionality limited to very small Excel worksheets.
- `Xlsxir.ConvertDate` functionality incomplete.
