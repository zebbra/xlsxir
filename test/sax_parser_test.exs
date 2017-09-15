defmodule SaxParserTest do
  use ExUnit.Case

  alias Xlsxir.SaxParser

  test "reads complex shared strings correctly" do
    {:ok, %{tid: tid}, _} = SaxParser.parse(File.read!("./test/test_data/complexStrings.xml"), :string)

    assert find_string(tid, 0) == "FOO: BAR"
    assert find_string(tid, 1) == "BAZ"
  end

  defp find_string(tid, index) do
    :ets.lookup(tid, index)
    |> List.first
    |> elem(1)
  end
end
