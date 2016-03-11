defmodule Xlsxir do

  alias Xlsxir.Unzip
  alias Xlsxir.Parse
  alias Xlsxir.Format

  @moduledoc """
  Parses and extracts data from a Microsoft Excel workbook.
  """

  @doc """
  Provides a list of row value lists -or- a map of cell/value pairs from a Microsoft Excel workbook
  given the path to a file of extension type `.xlsx`, the index of the worksheet in which the information 
  is requested, and an option.

  Cells containing formulas return a `string` of the resulting value. Cells containing Excel date format return
  a date in Erlang `date()` format (i.e. `{year, month, day}`).

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
    - cell 'D1' -> formula of `=4*5`
    - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

        iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
        [["string one", "string two", 10, 20, {2016, 1, 1}]]

        iex> Xlsxir.extract("./test/test_data/test.xlsx", 0, :cells)
        %{ A1: "string one", B1: "string two", C1: 10, D1: 20, E1: {2016,1,1}}
  """
  def extract(path, index, option \\ :rows) do
    strings = path
              |> Unzip.validate_path
              |> Parse.shared_strings

    path
    |> Unzip.validate_path
    |> Parse.worksheet(index)
    |> Format.prepare_output(strings, option)
  end
end
