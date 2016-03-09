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
    {:ok, "Good_Example.xlsx"}

    iex> path = "bad_path.xml"
    iex> Xlsxir.Unzip.validate_path(path)
    {:error, "Invalid path. Currently only .xlsx file types are supported."}
  """
  def validate_path(path) do
    path
    |> String.downcase
    |> String.split(".", trim: true)
    |> List.last
    |> case do
      "xlsx" -> {:ok, path}
      _      -> {:error, "Invalid path. Currently only .xlsx file types are supported."}
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
    |> :zip.zip_open([:memory])
    |> case do
      {:error, cause}      -> {:error, cause}
      {:ok, zip_directory} ->
        inner_path
        |> :zip.zip_get(zip_directory)
        |> case do
          {:error, cause}          -> {:error, cause}
          {:ok, {_, file_content}} ->
            case :zip.zip_close(zip_directory) do
              {:error, cause} -> {:error, cause}
              :ok             -> {:ok, file_content}
            end
        end
    end
  end

end