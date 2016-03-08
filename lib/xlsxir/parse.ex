defmodule Xlsxir.Parse do
  import Xlsxir.Unzip, only: [extract_xml: 2]
  import SweetXml

  @moduledoc """
    Receives Excel xml data via the `extract_xml` function of the `Unzip` module and parses it.
  """

  @doc """
    Receives Excel string data in xml format, parses it and returns the strings in a list.
  """
  def shared_strings(path) do
    {:ok, strings} = extract_xml(path, 'xl/sharedStrings.xml')
    strings
    |> xpath(~x"//t/text()"sl)
  end

  @doc """
    Receives the xlsx worksheet at position `index` in xml format, parses the data and returns
    required elements in the form of a map:

      %{ 'row_num' => %{ 'cell' => [ 's' for string or nil, '1' for date or nil,
        Excel function or nil, value or reference to sharedStrings ] }, ... }

    Example:

      %{ '1' =>                                 // row 1
        %{ 'A1' => [ 's', nil, nil, '0']},      // cell A1, string from position '0' of sharedStrings.xml
        %{ 'B1' => [ nil, nil, '4*5', '20']},   // cell B1, function with a value of '20'
        %{ 'C1' => [ nil, '1', nil, '41014']}   // cell C1, Excel date serial number of '41014'
      }
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