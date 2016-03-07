defmodule Xlsxir.Parse do
  import Xlsxir.Unzip, only: [extract_xml: 2]
  import SweetXml

  def shared_strings_with_index(path) do
    {:ok, strings} = extract_xml(path, 'xl/sharedStrings.xml')
    strings
    |> xpath(~x"//t/text()"sl)
    |> Enum.with_index
  end

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