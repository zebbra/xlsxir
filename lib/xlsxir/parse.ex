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
    {:ok, worksheet} = extract_xml(path, 'xl/worksheets/sheet#{index + 1}.xml')
    worksheet
    |> xmap(
        sheet: [
          ~x"//row"l,
          row: ~x"./@r",
          columns: [
            ~x"./c"l,
            column:   ~x"./@r",
            type:     ~x"./@t",
            function: ~x"./f/text()"o,
            value:    ~x"./v/text()"
            ] 
          ]
        )
  end
end