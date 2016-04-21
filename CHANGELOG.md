# Change Log

## 0.0.4

- Expanded coverage of Office Open XML standard `numFmt` (Standard Number Format). The `formatCode` for a standard `numFmt` is implied rather than explicitly identified in the XML file.
- Implemented support for Office Open XML custom `numFmt` (Custom Number Format) utilizing the `formatCode` explicitly identified in the XML file. 
- Added `Number Styles` documentation covering standard and custom `numFmt` and how to manually add an unsupported `numFmt`.

## 0.0.3

- Fixed issue related to strings that contain special characters. Refactored `Xlsxir.Parse.shared_strings/1` to properly parse strings with special characters.

## 0.0.2

- Refactored `Xlsxir.Parse` functions to improve `extract` performance on larger files.
- Expanded documentation and test coverage.
- Completed `Xlsxir.ConvertDate` module.

## 0.0.1

- Initial draft. Functionality limited to very small Excel worksheets.
- `Xlsxir.ConvertDate` functionality incomplete.
