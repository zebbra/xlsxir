defmodule Xlsxir.ParseString do
  alias Xlsxir.{SharedString, Index}

  @moduledoc """
  Holds the SAX event instructions for parsing sharedString data via `Xlsxir.SaxParser.parse/2`
  """

  defmodule StringState do
    defstruct empty_string: true
  end

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
  def sax_event_handler(:startDocument, _state), do: Index.new

  def sax_event_handler({:startElement,_,'t',_,_}, _state), do: %StringState{}

  def sax_event_handler({:characters, value}, state) do
    value
    |> to_string
    |> SharedString.add_shared_string(Index.get)

    Index.increment_1
    %{state | empty_string: false}
  end

  def sax_event_handler({:endElement,_,'t',_}, %StringState{empty_string: empty_string} = state) do
    if empty_string do 
      SharedString.add_shared_string("", Index.get)
      Index.increment_1
    end
  end

  def sax_event_handler(:endDocument, _state), do: Index.delete

  def sax_event_handler(_, _state), do: nil
 
end