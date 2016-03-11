defmodule DocTest do
  use ExUnit.Case
  doctest Xlsxir.Unzip
  doctest Xlsxir.Parse
  doctest Xlsxir.ConvertDate
  doctest Xlsxir.Format
  doctest Xlsxir
end