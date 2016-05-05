defmodule Xlsxir.Format do
  alias Xlsxir.{ConvertDate, Worksheet, SharedString}

  @moduledoc """
  Receives parsed excel worksheet data, formats the cell values and returns the data in either a
  list or a map depending on the option chosen.
  """

  @doc """
  Main function of `Format` module. Receives the parsed excel worksheet data, the parsed excel string
  data and the chosen format via `option` which defaults to `:rows`.
  """
  def prepare_output(option) do
    sheet = Worksheet.get 
    shared_strings = SharedString.get 

    case option do
      :rows  -> row_list(sheet, shared_strings)
      :cells -> cell_map(sheet, shared_strings)
      _      -> raise ArgumentError, message: "Invalid option."
    end
  end

  @doc """
  Formats parsed excel worksheet into a list of lists containing cell values by row.

  ## Parameters

  - `sheet` - map of xml worksheet data from Excel file that was parsed via Xlsxir.Parse.worksheet/2
  - `strings` - list of sharedStrings.xml from Excel file that was parsed via Xlsxir.Parse.shared_strings/1

  ## Example
      [[row_1_values], [row_2_values], ...]
  """
  def row_list(sheet, strings) do
    sheet
    |> Enum.chunk_by(fn cell -> Keyword.keys([cell])
                            |> List.first
                            |> Atom.to_string
                            |> regx_scan
                          end)
    |> Enum.map(fn row ->
        Enum.map(row, fn {_k, v} -> format_cell_value(v, strings) end)
      end)
  end

  @doc """
  Formats parsed excel worksheet into a map of cell/value pairs.

  ## Parameters

  - `sheet` - map of xml worksheet data from Excel file that was parsed via `Xlsxir.Parse.worksheet/2`
  - `strings` - list of sharedStrings.xml data from Excel file that was parsed via `Xlsxir.Parse.shared_strings/1`

  ## Example

      %{ A1: value_of_cell, B1: value_of_cell, ...}
  """
  def cell_map(sheet, strings) do
    sheet
    |> Enum.map(fn row ->
        Enum.map(row, fn{k, v} ->
          {k, format_cell_value(v, strings)}
        end)
      end)
    |> List.flatten
    |> Enum.reduce(%{}, &(Enum.into [&1], &2))
  end

  @doc """
  Uses attributes from xml to format cell value.

  ## Parameters

  - `list` - list containing attribute, style and value of column from xml file
  - `strings` - list of strings from the sharedStrings.xml file

  ## Example

      iex> Xlsxir.Format.format_cell_value([nil, nil, '1'], ["A", "B"])
      1
      iex> Xlsxir.Format.format_cell_value(['s', nil, '1'], ["A", "B"])
      "B"
  """
  def format_cell_value(list, strings) do
    case list do
      [ nil, nil, nil]  -> nil                                                                 # Empty cell without assigned attribute
      [   _,    _, ""]  -> ""                                                                  # Empty cell with assigned attribute
      [ 'e',  nil,  e]  -> List.to_string(e)                                                   # Excel type error
      [ 's',    _,  i]  -> Enum.at(strings, List.to_integer(i))                                # Excel type string
      [ nil,  nil,  n]  -> convert_char_number(n)                                              # Excel type number
      [ 'n',  nil,  n]  -> convert_char_number(n)
      [ nil,  'd',  d]  -> ConvertDate.from_excel(d)                                           # Excel type date
      [ 'n',  'd',  d]  -> ConvertDate.from_excel(d)
      ['str', nil,  s]  -> List.to_string(s)                                                   # Excel type formula w/ string 
      _                 -> raise "Unmapped attribute #{Enum.at(list, 0)}. Unable to process"   # Unmapped Excel type
    end
  end

  @doc """
  Converts Excel number to either integer or float.
  """
  def convert_char_number(number) do
    number
    |> List.to_string
    |> String.match?(~r/[.]/)
    |> case do
        false -> List.to_integer(number)
        true  -> List.to_float(number)
       end
  end

  defp regx_scan(cell) do
    ~r/[0-9]/
    |> Regex.scan(cell)
    |> List.to_string
  end

end
