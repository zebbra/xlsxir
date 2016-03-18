defmodule ParseTest do
  use ExUnit.Case
  doctest Xlsxir.Parse

  import Xlsxir.Parse

  test "second worksheet is parsed with index argument of 1" do
    assert worksheet("./test/test_data/test.xlsx", 1) == [[A1: [nil, '1'], B1: [nil, '2']], [A2: [nil, '3'], B2: [nil, '4']]]
  end

  test "able to parse maximum number of columns" do
    sheet3 = worksheet("./test/test_data/test.xlsx", 2)
    assert List.first(sheet3)[:XFD1] == [nil, '16384']
  end

  test "able to parse maximum number of rows" do
    sheet4 = worksheet("./test/test_data/test.xlsx", 3)
    assert List.last(sheet4)[:A1048576] == [nil, '1048576']
  end

end
