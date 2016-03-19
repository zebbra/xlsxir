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
  required elements in the form of a `keyword list`:

      [[row_1_cell_1: ['attribute', 'value'], ...], [row_2_cell_1: ['attribute', 'value'], ...], ...]

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
          [[A1: ['s', '0'], B1: ['s', '1'], C1: [nil, '10'], D1: [nil, '20'], E1: ['1', '42370']]]
  """
  def worksheet(path, index) do
    {:ok, sheet} = extract_xml(path, 'xl/worksheets/sheet#{index + 1}.xml')

    sheet
    |> xpath(~x"//worksheet/sheetData/row/c"l)
    |> Stream.map(&process_column/1)
    |> Enum.chunk_by(fn cell -> Keyword.keys([cell])
                                |> List.first
                                |> Atom.to_string
                                |> regx_scan
                              end)
  end

  defp process_column({:xmlElement,:c,:c,_,_,_,_,xml_attr,xml_elem,_,_,_}) do
    value = extract_value(xml_elem)
    {cell_ref, attribute} = extract_attribute(xml_attr)
    {List.to_atom(cell_ref), [attribute, value]}
  end

  defp extract_attribute(xml_attr) do
    n = Enum.count(xml_attr)

    cell_ref = case List.first(xml_attr) do
                 {:xmlAttribute, _,_,_,_,_,_,_,cell,_} -> cell
                 _                                     -> raise "Unassigned cell reference."
               end

    attribute = case List.last(xml_attr) do
                  {:xmlAttribute, _,_,_,_,_,_,_,attr,_} when n == 2 -> attr
                  _                                                 -> nil
                end

    {cell_ref, attribute}
  end

  defp extract_value(xml_elem) do
    case xml_elem do
      [{:xmlElement,_,_,_,_,_,_,_,[{_,_,_,_,val,_}],_,_,_}]         -> val
      [_,{:xmlElement,_,_,_,_,_,_,_,[{_,_,_,_,funct_val,_}],_,_,_}] -> funct_val
      []                                                            -> nil
      _                                                             -> raise "Invalid xmlElement."
    end
  end

  defp regx_scan(cell) do
    ~r/[0-9]/
    |> Regex.scan(cell)
    |> List.to_string
  end

end
