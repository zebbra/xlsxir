defmodule ParseTest do
  use ExUnit.Case
  doctest Xlsxir.Parse

  import Xlsxir.Parse

  def styles, do: []
  def custom_styles, do: num_style("./test/test_data/test.xlsx")

  test "second worksheet is parsed with index argument of 1" do
    assert worksheet("./test/test_data/test.xlsx", 1, styles) == [[A1: [nil, nil, '1'], B1: [nil, nil, '2']], [A2: [nil, nil, '3'], B2: [nil, nil, '4']]]
  end

  test "able to parse maximum number of columns" do
    sheet3 = worksheet("./test/test_data/test.xlsx", 2, styles)
    assert List.first(sheet3)[:XFD1] == [nil, nil, '16384']
  end

  test "able to parse maximum number of rows" do
    sheet4 = worksheet("./test/test_data/test.xlsx", 3, styles)
    assert List.last(sheet4)[:A1048576] == [nil, nil, '1048576']
  end

  test "able to parse cells with errors" do
    assert worksheet("./test/test_data/test.xlsx", 4, styles) == [[A1: ['e', nil, '#DIV/0!'], B1: ['e', nil, '#REF!'], C1: ['e', nil, '#NUM!'], D1: ['e', nil, '#VALUE!']]]
  end

  test "able to parse custom formats" do
    assert worksheet("./test/test_data/test.xlsx", 5, custom_styles) == [[A1: [nil, nil, '-123.45'], B1: [nil, nil, '67.89'], C1: [nil, 'd', '42005'], D1: [nil, 'd', '42735']]]
  end
end
