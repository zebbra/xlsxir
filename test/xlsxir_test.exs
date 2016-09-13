defmodule XlsxirTest do
  use ExUnit.Case

  import Xlsxir

  def path, do: "./test/test_data/test.xlsx"

  test "second worksheet is parsed with index argument of 1" do
    extract(path, 1)
    assert get_list == [[1, 2], [3, 4]]
    close
  end

  test "able to parse maximum number of columns" do
    extract(path, 2)
    assert get_cell("XFD1") == 16384
    close
  end

  test "able to parse maximum number of rows" do
    extract(path, 3)
    assert get_cell("A1048576") == 1048576
    close
  end

  test "able to parse cells with errors" do
    extract(path, 4)
    assert get_list == [["#DIV/0!", "#REF!", "#NUM!", "#VALUE!"]]
    close
  end

  test "able to parse custom formats" do
    extract(path, 5)
    assert get_list == [[-123.45, 67.89, {2015, 1, 1}, {2016, 12, 31}]]
    close
  end

  test "able to parse with conditional formatting" do
    extract(path, 6)
    close
  end
end
