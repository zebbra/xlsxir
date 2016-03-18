defmodule UnzipTest do
  use ExUnit.Case
  doctest Xlsxir.Unzip

  import Xlsxir.Unzip

  @path "./test/test_data/test.zip"
  @inner_path 'test.txt'
  @incorrect_path "./bad/path.zip"
  @incorrect_inner_path 'bad_inner_path.txt'

  test "path has the correct extension" do
    assert validate_path("correct_path.xlsx") == {:ok, 'correct_path.xlsx'}
  end

  test "path has incorrect extension" do
    assert validate_path("incorrect_path.xml") == {:error, "Invalid path. Currently only .xlsx file types are supported."}
  end

  # test.zip includes a sigle text file 'test.txt' which includes a single string "test_successful"
  test ".zip content extractable when correct path and inner_path given" do
    assert extract_xml(@path, @inner_path) == {:ok, "test_successful"}
  end

  test ".zip content extraction fails when incorrect path given" do
    assert extract_xml(@incorrect_path, @inner_path) == {:error, :enoent}
  end

  test ".zip content extraction fails when incorrect inner_path given" do
    assert extract_xml(@path, @incorrect_inner_path) == {:error, :file_not_found}
  end

end
