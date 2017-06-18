defmodule SaxParserTest do
  use ExUnit.Case

  alias Xlsxir.SaxParser
  alias Xlsxir.SharedString

  test "reads complex shared strings correctly" do
    SaxParser.parse("./test/test_data/complexStrings.xml", :string)

    assert SharedString.get_at(0) == "FOO: BAR"
    assert SharedString.get_at(1) == "BAZ"
  end
end
