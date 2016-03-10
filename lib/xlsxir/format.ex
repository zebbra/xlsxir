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
    |> Enum.map(fn{k,v} -> 
        Enum.map(v, fn{k2, v2} -> 
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
    
  end

  # type string
  defp format_cell_value(list = ['s', nil, nil, _], strings) do
    [_, _, _, n] = list
    Enum.at(strings, List.to_integer(n))
  end

  # type integer
  defp format_cell_value(list = [nil, nil, nil, _], _) do
    [_, _, _, i] = list
    List.to_float(i)
  end

  # type date
  defp format_cell_value(list = [nil, '1', nil, _], _) do
    [_, _, _, date_serial] = list
    Xlsxir.ConvertDate.from_excel(date_serial)
  end

  # type formula
  defp format_cell_value(list = [nil, nil, _, _], _) do
    [_, _, _, value] = list
    List.to_string(value)
  end
  
  # no match
  defp format_cell_value(_, _) do
    raise "Data corrupt. Unable to process."
  end

end





