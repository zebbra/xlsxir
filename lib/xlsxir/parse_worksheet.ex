defmodule Xlsxir.ParseWorksheet do
  alias Xlsxir.{Worksheet, SharedString, Style, Index, ConvertDate}
  import Xlsxir.ConvertDate, only: [convert_char_number: 1]
  require IEx

  @moduledoc """
  Holds the SAX event instructions for parsing worksheet data via `Xlsxir.SaxParser.parse/2`
  """
  
  defmodule RowState do
    defstruct row: %{}, cell_ref: "", data_type: "", num_style: "", value: ""
  end
  
  @doc """
  Sax event utilized by `Xlsxir.SaxParser.parse/2`. Takes a pattern and the current state of a struct and recursivly parses the
  worksheet XML file, ultimately sending a keyword list of cell references and their assocated data to the `Worksheet` agent 
  process that was started by `Xlsxir.SaxParser.parse/2`. 

  ## Parameters

  - pattern - the XML pattern of the event to match upon
  - state - the state of the `%RowState{}` struct which temporarily holds applicable data of the current row being parsed

  ## Example
  Each entry in the keyword list created consists of a cell reference atom and a list containing the cell's data type, style index
  and value (i.e. `[A1: ['s', nil, '0'], ...]`).
  """

  def sax_event_handler(:startDocument, _state), do: Index.new

  def sax_event_handler({:startElement,_,'row',_,_}, _state), do: state = %RowState{}

  def sax_event_handler({:startElement,_,'c',_,xml_attr}, state) do
    a = Enum.map(xml_attr, fn(attr) -> 
      case attr do
        {:attribute,'r',_,_,ref}   -> {:r, ref  }
        {:attribute,'s',_,_,style} -> {:s, if Style.alive? do
                                             Style.get_at(List.to_integer(style))
                                           else 
                                             nil
                                           end}
        {:attribute,'t',_,_,type}  -> {:t, type }
        _                                      -> raise "Unknown cell attribute"
      end
    end)

    {cell_ref, num_style, data_type} = case Keyword.keys(a) do
                                         [:r]         -> {a[:r],   nil,   nil}
                                         [:s, :r]     -> {a[:r], a[:s],   nil}
                                         [:t, :r]     -> {a[:r],   nil, a[:t]}
                                         [:t, :s, :r] -> {a[:r], a[:s], a[:t]}
                                         _            -> raise "Invalid attributes: #{a}"
                                       end
    %{state | cell_ref: cell_ref, num_style: num_style, data_type: data_type}
  end

  def sax_event_handler({:characters, value}, state) do
    if state == nil, do: nil, else: %{state | value: value}
  end

  def sax_event_handler({:endElement,_,'c',_}, %RowState{row: row} = state) do
    if state.data_type == 's', do: IO.inspect([state.data_type, state.num_style, state.value])
    cell_value = format_cell_value([state.data_type, state.num_style, state.value])


    %{state | row: Enum.into(row, [[to_string(state.cell_ref), cell_value]])} 
  end

  def sax_event_handler({:endElement,_,'row',_}, state) do
    state.row
    |> Enum.reverse
    |> Worksheet.add_row(Index.get)

    Index.increment_1
  end

  def sax_event_handler(:endDocument, _state) do 
    Index.delete
    SharedString.delete
    Style.delete
  end

  def sax_event_handler(_, state), do: state

  defp format_cell_value(list) do
    case list do
      [ nil, nil, nil]  -> nil                                                                 # Empty cell without assigned attribute
      [   _,    _, ""]  -> ""                                                                  # Empty cell with assigned attribute
      [ 'e',  nil,  e]  -> List.to_string(e)                                                   # Excel type error
      [ 's',    _,  i]  -> SharedString.get_at(List.to_integer(i))                             # Excel type string
      [ nil,  nil,  n]  -> convert_char_number(n)                                              # Excel type number
      [ 'n',  nil,  n]  -> convert_char_number(n)
      [ nil,  'd',  d]  -> ConvertDate.from_excel(d)                                           # Excel type date
      [ 'n',  'd',  d]  -> ConvertDate.from_excel(d)
      ['str', nil,  s]  -> List.to_string(s)                                                   # Excel type formula w/ string 
      _                 -> raise "Unmapped attribute #{Enum.at(list, 0)}. Unable to process"   # Unmapped Excel type
    end
  end

end