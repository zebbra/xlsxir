# Number Styles

When a cell containing a number has a format applied to it in a `.xlsx` document, a `<xf>` element is assigned to that cell. These elements are defined in the `xl/styles.xml` document of a `.xlsx` workbook. `numFmtId` is one of the attributes of the `<xf>` element. There are two general types of `numFmtId`s which, for purposes of this document, will be referred to as `standard` and `custom`. 

When a format is applied to a number, the actual value of the number does not change. Instead, a `formatCode` is applied to the cell which defines how the value contained within the cell should be displayed. The `numFmtId` is actually a reference to a `formatCode`. 

Excel has many baked-in number formats. The `formatCode` for these formats are implied rather than explicitly stated in the `xl/styles.xml` file and are referenced by `standard numFmtId`s. Custom formats on the other hand can vary from system to system and are therefore explicitly defined in the `xl/styles.xml` file. `formatCode`s for custom formats are referenced by `custom numFmtId`s. 

Xlsxir has been designed to capture the majority of both standard and custom formats, however, due to the number of varying types there is a possibility a few could have slipped through. If you receive an error of `Unsupported style type: x`, that is most likely the issue. The "style type" is simply a `numFmtId` that is currently not supported by Xlsxir. Any time unsupported style types are identified, please submit an [issue](https://github.com/kennellroxco/xlsxir/issues) on GitHub. If you would like to manually add the style type to your local instance of Xlsxir, only a few simple steps are required as outlined below. 

The `Xlsxir.ParseStyle` module handles the parsing of the `xl/styles.xml` file. It is designed to determine whether the underlying value of each `formatCode` is either a number (which includes both `integer` and `float`) or a date serial number and return a list of `numFmtId`s which have been translated to either `nil` for numbers or `'d'` for dates. 

To manually update the code for an unsupported style type, locate the `Xlsxir.ParseStyle` module. The top portion of the module should look like this: 

```
defmodule Xlsxir.ParseStyle do
  alias Xlsxir.{Style, Index}

  @moduledoc """
  Holds the SAX event instructions for parsing style data via `Xlsxir.SaxParser.parse/2`
  """

  # the following module attributes hold `numStyleId`s for standard number styles, grouping them between numbers and dates
  @num  [0,1,2,3,4,9,10,11,12,13,37,38,39,40,44,48,49,59,60,61,62,67,68,69,70]
  @date [14,15,16,17,18,19,20,21,22,27,30,36,45,46,47,50,57]
```

The module attributes `@num` and `@date` are lists of `standard numFmtId`s. Determine whether the underlying value of your unsupported "style type" is a number or a date and then add the "style type" to the appropriate module attribute list. 

If you have any trouble with this, please submit an [issue](https://github.com/kennellroxco/xlsxir/issues).
