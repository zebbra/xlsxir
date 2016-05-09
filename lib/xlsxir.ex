defmodule Xlsxir do
  require IEx
  alias Xlsxir.{Unzip, SaxParser, Worksheet, Timer}

  @moduledoc """
  Extracts and parses data from a Microsoft Excel workbook, returning either a list or a map.
  """

  @doc """
  Provides a list of row value lists or a map of cell/value pairs from a Microsoft Excel workbook
  given the path to a file of extension type `.xlsx`, the index of the worksheet in which the information
  is requested, and an option.

  Cells containing formulas return either a `string`, `integer` or `float` of the resulting value. Cells containing Excel date format return
  a date in Erlang `:calendar.date()` format (i.e. `{year, month, day}`).

  ## Parameters
  - `path` - file path of a `.xlsx` file type in `string` format
  - `index` - index of worksheet from within the Excel workbook to be parsed, starting with `0`
  - `option` - one of two available options

  ## Options

  - `:rows` - a list of row value lists (default) - i.e. [[row_1_values], [row_2_values], ...]
  - `:cells` - a map of cell/value pairs - i.e. %{ A1: value_of_cell, B1: value_of_cell, ...}

  ## Example
    An example file named `test.xlsx` located in `./test/test_data` containing the following:
    - cell 'A1' -> "string one"
    - cell 'B1' -> "string two"
    - cell 'C1' -> integer of 10
    - cell 'D1' -> formula of "4 * 5"
    - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
          [["string one", "string two", 10, 20, {2016, 1, 1}]]

          iex> Xlsxir.extract("./test/test_data/test.xlsx", 0, :cells)
          %{ A1: "string one", B1: "string two", C1: 10, D1: 20, E1: {2016,1,1}}
  """
  def extract(path, index) do
    Timer.start
    {:ok, file}       = Unzip.validate_path(path)
    {:ok, file_paths} = Unzip.xml_file_list(index)
                        |> Unzip.extract_xml_to_file(file)

    Enum.each(file_paths, fn file -> 
      case file do
        'temp/xl/sharedStrings.xml' -> SaxParser.parse(to_string(file), :string)
        'temp/xl/styles.xml'        -> SaxParser.parse(to_string(file), :style)
        _                           -> SaxParser.parse(to_string(file), :worksheet)
      end
    end)

    Unzip.delete_temp_dir
    {:ok, Timer.stop}
  end

  def get_list do
    range = 0..(:ets.info(:worksheet, :size) -1)

    range
    |> Enum.map(fn i -> Worksheet.get_at(i)
                        |> Enum.map(fn cell -> Enum.at(cell, 1) end)
                      end)
  end

  def get_map do
    range = 0..(:ets.info(:worksheet, :size) -1)

    range
    |> Enum.reduce(%{}, fn i, m -> Worksheet.get_at(i)
                                   |> Enum.reduce(%{}, fn [k,v], acc -> Map.put(acc, k, v) end)
                                   |> Enum.into(m)
                                 end)
  end

  def get_cell(cell_ref) do
    default = "No value found in cell #{cell_ref}"

    row = ~r/\d+/
          |> Regex.scan(cell_ref)
          |> List.flatten
          |> List.first
          |> String.to_integer

    [_, value] = Enum.find(Worksheet.get_at(row - 1), default, fn cell -> 
                   Enum.at(cell, 0) == cell_ref 
                 end)
    value
  end

  def get_row(row) do
    Enum.map(Worksheet.get_at(row - 1), fn [k,v] -> v end)
  end

  def get_col(col) do
    Enum.map(get_map, fn {k,v} -> if cell_ltrs(k) == col, do: v end)
    |> Enum.reject(fn x -> x == nil end)
  end

  def close do
    Worksheet.delete
  end

  defp cell_ltrs(cell) do
    ~r/[a-z]+/i 
    |> Regex.scan(String.upcase(cell))
    |> List.flatten
    |> List.first
  end

end
