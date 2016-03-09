defmodule Xlsxir.Parse do
  import Xlsxir.Unzip, only: [extract_xml: 2]
  import SweetXml

  @moduledoc """
  Receives Excel xml data via the `extract_xml` function of the `Unzip` module and parses it.
  """

  @doc """
  Receives Excel string data in xml format, parses it and returns the strings in the form of a list.

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format

  ## Example
    An example file named `test.xlsx` located in `./test/test_data` containing the following:
    - cell 'A1' -> "string one"
    - cell 'B1' -> "string two"
    - cell 'C1' -> integer of 10
    - cell 'D1' -> formula of `=4*5`
    - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

      iex> Xlsxir.Parse.shared_strings("./test/test_data/test.xlsx")
      ["string one", "string two"]
  """
  def shared_strings(path) do
    {:ok, strings} = extract_xml(path, 'xl/sharedStrings.xml')
    strings
    |> xpath(~x"//t/text()"sl)
  end

  @doc """
  Receives the xlsx worksheet at position `index` in xml format, parses the data and returns
  required elements in the form of a `map` of `char_lists`:

      %{ 'row_num' => %{ 'cell' => [ 's' for string or nil, '1' for date or nil,
        Excel function or nil, value or reference to sharedStrings ] }, ...}

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format
  - `index` - index of worksheet from within the Excel workbook to be parsed, starting with `0`

  ## Example
    An example file named `test.xlsx` located in `./test/test_data` containing the following in worksheet at index `0`:
    - cell 'A1' -> "string one"
    - cell 'B1' -> "string two"
    - cell 'C1' -> integer of 10
    - cell 'D1' -> formula of `=4*5`
    - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

      iex> Xlsxir.Parse.worksheet("./test/test_data/test.xlsx", 0)
      %{'1' => %{'A1' => ['s', nil, nil, '0'], 'B1' => ['s', nil, nil, '1'], 'C1' => [nil, nil, nil, '10'], 
        'D1' => [nil, nil, '4*5', '20'], 'E1' => [nil, '1', nil, '42370']}}
  """
  def worksheet(path, index) do
    {:ok, sheet} = extract_xml(path, 'xl/worksheets/sheet#{index + 1}.xml')

    sheet
    |> xpath(~x"//row/@r"l)
    |> Enum.reduce(%{}, fn(row, acc) -> Map.put(acc, row,
      sheet
      |> xpath(~x"//row[contains(@r, '#{row}')]/c/@r"l)
      |> Enum.reduce(%{}, fn(cell, acc) -> Map.put(acc, cell,
          [      
            col_data(sheet, cell, "@t"),
            col_data(sheet, cell, "@s"),
            col_data(sheet, cell, "f/text()", "o"),
            col_data(sheet, cell, "v/text()")
          ])
        end)
      )end)
  end

  defp col_data(xml, cell, data, opt \\ "") do
    xpath(xml, ~x"//c[contains(@r, '#{cell}')]/#{data}"opt)
  end  
end