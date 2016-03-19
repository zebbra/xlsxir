defmodule Xlsxir.Unzip do

  @moduledoc """
  Provides validation of accepted file extension types for file path and extracts requested inner files
  from a `.zip` file type.
  """

  @doc """
  Validates given path is of extension type `.xlsx` and returns a tuple.

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format

  ## Example

      iex> path = "Good_Example.xlsx"
      iex> Xlsxir.Unzip.validate_path(path)
      {:ok, 'Good_Example.xlsx'}

      iex> path = "bad_path.xml"
      iex> Xlsxir.Unzip.validate_path(path)
      {:error, "Invalid path. Currently only .xlsx file types are supported."}
  """
  def validate_path(path) do
    path = to_string path
    cond do
      Regex.match?(~r/\.xlsx$/, path) -> {:ok, to_char_list path}
      true -> {:error, "Invalid path. Currently only .xlsx file types are supported."}
    end
  end

  @doc """
  Extracts requested file from a `.zip` file type and returns a tuple.

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format
  - `inner_path` - file path from within a `.zip` file type of the document being requested in `char_list` format

  ## Example

    An example file named `test.zip` located in "./test_data/test" containing a single file named `test.txt`
    containing a single string of "test_successful":

        iex> path = "./test/test_data/test.zip"
        iex> inner_path = 'test.txt'
        iex> Xlsxir.Unzip.extract_xml(path, inner_path)
        {:ok, "test_successful"}
  """

  def extract_xml(path, inner_path) do
    path
    |> to_char_list
    |> :zip.extract([:memory, {:file_filter, fn(file) -> elem(file, 1) == inner_path end}])
    |> case do
        {:ok, [{_, file_content}]} -> {:ok, file_content}
        {:ok, []}                  -> {:error, :file_not_found}
        {:error, cause}            -> {:error, cause}
       end
  end

end
