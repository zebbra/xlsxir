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
    |> Enum.map(fn{_, v} -> 
        Enum.map(v, fn{_, v2} -> 
          format_cell_value(v2, strings) 
        end) 
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
    Map.values(sheet)
    |> Enum.map(fn row -> 
        Enum.map(row, fn{k, v} -> 
          {List.to_atom(k), format_cell_value(v, strings)} 
        end) 
      end)
    |> List.flatten
    |> Enum.reduce(%{}, &(Enum.into [&1], &2))
  end

  defp format_cell_value(list, strings) do
    case list do
      ['s', nil, nil, i]   -> Enum.at(strings, List.to_integer(i))       # Excel type string
      [nil, nil, nil, n]   -> convert_char_number(n)                     # Excel type number
      [nil, '1', nil, d]   -> Xlsxir.ConvertDate.from_excel(d)           # Excel type date
      [nil, nil,   _, f_n] -> convert_char_number(f_n)                   # Excel type formula w/ number
      ['str', nil, _, f_s] -> List.to_string(f_s)                        # Excel type Formula w/ string
      _                    -> raise "Data corrupt. Unable to process"    # invalid Excel type
    end
  end

  defp convert_char_number(number) do
    number
    |> List.to_string
    |> String.match?(~r/[.]/)
    |> case do
        false -> List.to_integer(number)
        true  -> List.to_float(number)
       end
  end

  def col_letter(i), do: do_col_letter(i, [])

  def do_col_letter(i, ltrs) when i/26 >= 1 do
    ltr = rem(i, 26) + 65

    i/26 - 1
    |> Float.floor
    |> round
    |> do_col_letter([ltr|ltrs])
  end

  def do_col_letter(i, ltrs) do
    ltr = rem(i, 26) + 65

    [ltr|ltrs]
    |> Enum.map(fn(x) -> <<x>> end)
    |> List.to_string
  end

end





