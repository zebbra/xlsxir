defmodule FormatTest do
  use ExUnit.Case
  doctest Xlsxir.Format

  import Xlsxir.Format

  test "a string is returned when cell value is a string" do
    from_cell = format_cell_value(['s', nil, nil, '0'], ["A", "B"])
    assert is_binary(from_cell)
  end

  test "when the cell contains an Excel function, a string of the
    resulting value is returned" do
      assert format_cell_value([nil, nil, '3*3', '9'], ["A", "B"]) == "9"
    end

  test "when column attributes are nil, return an integer" do
    assert format_cell_value([nil, nil, nil, '5'], ["A", "B"]) == 5
  end

end