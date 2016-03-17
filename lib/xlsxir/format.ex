defmodule Xlsxir.Format do
  
  @moduledoc """
  Receives parsed excel worksheet data, formats the cell values and returns the data in either a
  list or a map depending on the option chosen. 
  """

  @doc """
  Main function of `Format` module. Receives the parsed excel worksheet data, the parsed excel string
  data and the chosen format via `option` which defaults to `:rows`.
  """
  def prepare_output(worksheet, shared_strings, option) do
    case option do
      :rows  -> row_list(worksheet, shared_strings)
      :cells -> cell_map(worksheet, shared_strings)
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

  - `list` - list containing attribute and value of column from xml file
  - `strings` - list of strings from sharedStrings.xml file

  ## Example

      iex> Xlsxir.Format.format_cell_value([nil, '1'], ["A", "B"])
      1
      iex> Xlsxir.Format.format_cell_value(['s', '1'], ["A", "B"])
      "B"
  """
  def format_cell_value(list, strings) do
    case list do
      ['s', i]     -> Enum.at(strings, List.to_integer(i))                    # Excel type string
      [nil, n]     -> convert_char_number(n)                                  # Excel type number
      ['1', d]     -> Xlsxir.ConvertDate.from_excel(d)                        # Excel type date
      ['str', f_s] -> List.to_string(f_s)                                     # Excel type Formula w/ string
      _            -> raise "Invalid attribute #{list}. Unable to process"    # invalid Excel type
    end
  end

  # convert Excel number to either integer or float
  def convert_char_number(number) do
    number
    |> List.to_string
    |> String.match?(~r/[.]/)
    |> case do
        false -> List.to_integer(number)
        true  -> List.to_float(number)
       end
  end


#### Save for future development ####

#   # generate reference list for all cells
#   defp cell_reference_list do
#     Stream.flat_map(1..1048576, fn n -> 
#       Stream.map(0..16384, fn i -> String.to_atom(col_letter(i) <> Integer.to_string(n)) end)
#     end)
#   end

#   defp row_reference_list(n) do
#     Enum.map(0..16384, fn i -> String.to_atom(col_letter(i) <> Integer.to_string(n)) end)
#   end

#   # given index, return Excel column letter (i.e. 0 -> "A", 26 -> "AA")
#   defp col_letter(i), do: do_col_letter(i, [])

#   defp do_col_letter(i, ltrs) when i/26 >= 1 do
#     ltr = rem(i, 26) + 65

#     i/26 - 1
#     |> Float.floor
#     |> round
#     |> do_col_letter([ltr|ltrs])
#   end

#   defp do_col_letter(i, ltrs) do
#     ltr = rem(i, 26) + 65

#     [ltr|ltrs]
#     |> Enum.map(fn(x) -> <<x>> end)
#     |> List.to_string
#   end

end





