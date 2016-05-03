alias Xlsxir.{Worksheet, Style, String} 

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
  
end

defmodule Xlsxir.ParseStyle do
  
end
