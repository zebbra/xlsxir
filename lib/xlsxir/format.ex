defmodule Xlsxir.Format do
  
  @moduledoc """
    Receives parsed excel worksheet data, formats the cell values and returns the data in either a
    list or a map depending on the option chosen. 
  """

  @doc """
    Main call of `Format` module. Receives the parsed excel worksheet data, the parsed excel string
    data and the chosen format via `option` which defaults to `rows`.
  """
  def do_format(worksheet, shared_strings, option \\ 'rows') do
    case option do
      'rows'  -> row_list(worksheet, shared_strings)
      'cells' -> cell_map(worksheet, shared_strings)
      _       -> raise ArgumentError, message: "Invalid option."
    end
  end

  @doc """
    Formats parsed excel worksheet into a list of lists containing cell values by 
    row (i.e. [[row_1_values], [row_2_values], ...]).
  """
  def row_list(sheet, strings) do
    sheet
    |> Enum.map(fn{k,v} -> 
        Enum.map(v, fn{k2, v2} -> 
          #format_cell_value(v2, strings) 
          IO.inspect(v2)
          IO.puts("\n")
        end) 
      end)
  end

  @doc """
    Formats parsed excel worksheet into a map of cell/value pairs (i.e. %{"A1" => 
    value_of_cell, ...}).
  """
  def cell_map(sheet, strings) do
    
  end

  @doc """
    Formats the value of the string based upon its content.
  """
  @spec format_cell_value(list, list) :: String.t | integer
  # def format_cell_value(list, strings) do
  #   cond do  
  #     ['s', nil, nil, n] = list           -> Enum.at(strings, List.to_integer(n))
  #     [nil, nil, nil, i] = list           -> List.to_integer(i)
  #     [nil, nil, _, value] = list         -> List.to_string(value)
  #     [nil, '1', nil, date_serial] = list -> "date" #Xlsxir.ConvertDate.from_excel(date_serial)
  #     true                                -> raise "Data corrupt. Unable to process."
  #   end
  # end

  # type string
  def format_cell_value(list = ['s', nil, nil, _], strings) do
    [_, _, _, n] = list
    Enum.at(strings, List.to_integer(n))
  end

  # type integer
  def format_cell_value(list = [nil, nil, nil, _], _) do
    [_, _, _, i] = list
    List.to_integer(i)
  end

  # type date
  def format_cell_value(list = [nil, '1', nil, _], _) do
    [_, _, _, date_serial] = list
    "date" #Xlsxir.ConvertDate.from_excel(date_serial)
  end

  # type formula
  def format_cell_value(list = [nil, nil, _, _], _) do
    [_, _, _, value] = list
    List.to_string(value)
  end
  
  # no match
  def format_cell_value(_, _) do
    raise "Data corrupt. Unable to process."
  end

end





