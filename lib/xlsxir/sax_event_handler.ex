alias Xlsxir.{Worksheet, Style, SharedString} 

defmodule Xlsxir.ParseWorksheet do
  
  defmodule CellState do
    defstruct cell_ref: "", data_type: "", num_style: "", value: ""
  end
  
  def sax_event_handler({:startElement,_,'c',_,xml_attr}, _state) do
    state = %CellState{}
    
    a = Enum.map(xml_attr, fn(attr) -> 
      case attr do
        {:attribute,'r',_,_,ref}   -> {:r, ref  }
        {:attribute,'s',_,_,style} -> {:s, style}
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
    %{state | value: value}
  end

  def sax_event_handler({:endElement,_,'c',_}, state) do
    Worksheet.add_cell(List.to_atom(state.cell_ref), [state.data_type, state.num_style, state.value]) 
  end

  def sax_event_handler(:endDocument, state), do: state
  def sax_event_handler(_, state), do: state

end

defmodule Xlsxir.ParseString do

  def sax_event_handler({:characters, value}, _state) do
    value
    |> to_string
    |> SharedString.add_shared_string
  end

  def sax_event_handler(_, state), do: state
 
end

defmodule Xlsxir.ParseStyle do

  @num  [0,1,2,3,4,9,10,11,12,13,37,38,39,40,44,48,49,59,60,61,62,67,68,69,70]
  @date [14,15,16,17,18,19,20,21,22,27,30,36,45,46,47,50,57]

  defmodule CustomState do
    defstruct custom: %{}
  end

  def sax_event_handler(:startDocument, _state), do: state = %CustomState{}

  def sax_event_handler({:startElement,_,'xf',_,xml_attr}, state) do
    [{_,_,_,_,id}] = Enum.filter(xml_attr, fn attr -> case attr do 
                                                        {:attribute,'numFmtId',_,_,_} -> true
                                                        _                             -> false
                                                      end  
                                                    end)

    Style.add_id(id)
    state
  end

  def sax_event_handler({:startElement,_,'numFmt',_,xml_attr}, %CustomState{custom: custom} = state) do
    temp = Enum.reduce(xml_attr, %{}, fn attr, acc -> case attr do
                                                        {:attribute,'numFmtId',_,_,id}   -> Map.put(acc, :id, id)
                                                        {:attribute,'formatCode',_,_,cd} -> Map.put(acc, :cd, cd)
                                                        _                                -> nil
                                                      end
                                                    end)
    %{state | custom: Map.put(custom, temp[:id], temp[:cd])}
  end

  def sax_event_handler(:endDocument, %CustomState{custom: custom} = state) do
    custom_type = custom_handler(custom)

    Enum.each(Style.get_id, fn style_type -> 
      case List.to_integer(style_type) do
        i when i in @num   -> Style.add_style(nil)
        i when i in @date  -> Style.add_style('d')
        _                  -> if Map.has_key?(custom_type, style_type) do
                                Style.add_style(custom_type[style_type])
                              else
                                raise "Unsupported style type: #{style_type}. See doc page \"Number Styles\" for more info."
                              end
      end
    end)

    Style.delete_id
  end

  def sax_event_handler(_, state), do: state

  defp custom_handler(custom) do
    custom
    |> Enum.reduce(%{}, fn {k, v}, acc -> 
         cond do
           String.match?(to_string(v), ~r/\bred\b/i) -> Map.put_new(acc, k, nil)
           String.match?(to_string(v), ~r/[dhmsy]/i) -> Map.put_new(acc, k, 'd')
           true                                      -> Map.put_new(acc, k, nil)
         end
      end)
  end

end
