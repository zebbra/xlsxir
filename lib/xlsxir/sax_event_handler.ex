alias Xlsxir.{Worksheet, Style, SharedString} 

defmodule Xlsxir.ParseWorksheet do
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
  - state - the state of the `%CellState{}` struct which temporarily holds applicable data of the current cell being parsed

  ## Example
  Each entry in the keyword list created consists of a cell reference atom and a list containing the cell's data type, style index
  and value (i.e. `[A1: ['s', nil, '0'], ...]`).
  """

  def sax_event_handler({:startElement,_,'row',_,_}, _state) do
    state = %RowState{}
  end

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
    %{state | row: Enum.into(row, [[to_string(state.cell_ref), [state.data_type, state.num_style, state.value]]])} 
  end

  def sax_event_handler({:endElement,_,'row',_}, state) do
    state.row
    |> Enum.reverse
    |> Worksheet.add_row
  end

  def sax_event_handler(_, state), do: state

end

defmodule Xlsxir.ParseString do
  @moduledoc """
  Holds the SAX event instructions for parsing sharedString data via `Xlsxir.SaxParser.parse/2`
  """

  @doc """
  Sax event utilized by `Xlsxir.SaxParser.parse/2`. Takes a pattern and the current state of a struct and recursivly parses the
  sharedString XML file, ultimately sending each parsed string to the `SharedString` agent process that was started by 
  `Xlsxir.SaxParser.parse/2`. 

  ## Parameters

  - pattern - the XML pattern of the event to match upon
  - state - the state argument is unused when parsing strings and is therefore proceeded by an underscore

  ## Example
  Recursively sends strings from the `xl/sharedStrings.xml` file to `SharedString.add_shared_string/1`. The data can ultimately
  be retreived by the `get/0` function of the agent process (i.e. `Xlsxir.SharedString.get` would return `["string 1", "string 2", ...]`).
  """
  def sax_event_handler({:characters, value}, _state) do
    value
    |> to_string
    |> SharedString.add_shared_string
  end

  def sax_event_handler(_, _state), do: nil
 
end

defmodule Xlsxir.ParseStyle do
  @moduledoc """
  Holds the SAX event instructions for parsing style data via `Xlsxir.SaxParser.parse/2`
  """

  @num  [0,1,2,3,4,9,10,11,12,13,37,38,39,40,44,48,49,59,60,61,62,67,68,69,70]
  @date [14,15,16,17,18,19,20,21,22,27,30,36,45,46,47,50,57]

  defmodule CustomStyleState do
    defstruct custom_style: %{}
  end

  @doc """
  Sax event utilized by `Xlsxir.SaxParser.parse/2`. Takes a pattern and the current state of a struct and recursivly parses the
  styles XML file, ultimately sending each parsed style type to the `Style` agent process that was started by 
  `Xlsxir.SaxParser.parse/2`. The style types generated are `nil` for numbers and `'d'` for dates. 

  ## Parameters

  - pattern - the XML pattern of the event to match upon
  - state - the state of the `%CustomStyleState{}` struct which temporarily holds each `numFmtId` and its associated `formatCode` for custom format types

  ## Example
  Recursively sends style types generated from parsing the `xl/sharedStrings.xml` file to `Style.add/1`. The data can ultimately
  be retreived by the `get/0` function of the agent process (i.e. `Xlsxir.Style.get` would return something like `[nil, 'd', ...]` depending on
  each style type generated).
  """
  def sax_event_handler(:startDocument, _state), do: state = %CustomStyleState{}

  def sax_event_handler({:startElement,_,'xf',_,xml_attr}, state) do
    [{_,_,_,_,id}] = Enum.filter(xml_attr, fn attr -> 
                       case attr do 
                         {:attribute,'numFmtId',_,_,_} -> true
                         _                             -> false
                       end  
                     end)

    Style.add_id(id)
    state
  end

  def sax_event_handler({:startElement,_,'numFmt',_,xml_attr}, 
    %CustomStyleState{custom_style: custom_style} = state) do
    
    temp = Enum.reduce(xml_attr, %{}, fn attr, acc -> 
            case attr do
              {:attribute,'numFmtId',_,_,id}   -> Map.put(acc, :id, id)
              {:attribute,'formatCode',_,_,cd} -> Map.put(acc, :cd, cd)
              _                                -> nil
            end
          end)

    %{state | custom_style: Map.put(custom_style, temp[:id], temp[:cd])}
  end

  def sax_event_handler(:endDocument, 
    %CustomStyleState{custom_style: custom_style} = state) do

    custom_type = custom_style_handler(custom_style)

    Enum.each(Style.get_id, fn style_type -> 
      case List.to_integer(style_type) do
        i when i in @num   -> Style.add(nil)
        i when i in @date  -> Style.add('d')
        _                  -> if Map.has_key?(custom_type, style_type) do
                                Style.add(custom_type[style_type])
                              else
                                raise "Unsupported style type: #{style_type}. 
                                  See doc page \"Number Styles\" for more info."
                              end
      end
    end)

    Style.delete_id
  end

  def sax_event_handler(_, state), do: state

  defp custom_style_handler(custom_style) do
    custom_style
    |> Enum.reduce(%{}, fn {k, v}, acc -> 
         cond do
           String.match?(to_string(v), ~r/\bred\b/i) -> Map.put_new(acc, k, nil)
           String.match?(to_string(v), ~r/[dhmsy]/i) -> Map.put_new(acc, k, 'd')
           true                                      -> Map.put_new(acc, k, nil)
         end
      end)
  end

end
