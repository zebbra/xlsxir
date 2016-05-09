defmodule Xlsxir.Unzip do

  @moduledoc """
  Provides validation of accepted file extension types for file path, extracts required `.xlsx` contents to `./temp` and ultimately deletes the `./temp` directory and its contents.
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
  List of files contained in the requested `.xlsx` file to be extracted.

  ## Parameters

  - `index` - index of the `.xlsx` worksheet to be parsed

  ## Example

         iex> Xlsxir.Unzip.xml_file_list(0)
         ['xl/styles.xml', 'xl/sharedStrings.xml', 'xl/worksheets/sheet1.xml']
  """
  def xml_file_list(index) do
    [
     'xl/styles.xml',
     'xl/sharedStrings.xml',
     'xl/worksheets/sheet#{index + 1}.xml'
    ]
  end

  @doc """
  Extracts requested list of files from a `.zip` file to `./temp` and returns a list of the extracted file paths.

  ## Parameters

  - `file_list` - list containing file paths to be extracted in `char_list` format
  - `path` - file path of a `.xlsx` file type in `string` format

  ## Example
  An example file named `test.zip` located in './test_data/test' containing a single file named `test.txt`:

      iex> path = "./test/test_data/test.zip"
      iex> file_list = ['test.txt']
      iex> Xlsxir.Unzip.extract_xml_to_file(file_list, path)
      {:ok, ['temp/test.txt']}
      iex> File.ls("./temp")
      {:ok, ["test.txt"]}
      iex> Xlsxir.Unzip.delete_dir(["./temp"])
      :ok
  """
  def extract_xml_to_file(file_list, path) do
    path
    |> to_char_list
    |> :zip.extract([{:file_list, file_list}, {:cwd, 'temp/'}])
    |> case do 
        {:error, cause}   -> {:error, cause}
        {:ok, []}         -> {:error, :file_not_found}
        {:ok, files_list} -> {:ok, files_list}
       end
  end

  @doc """
  Deletes all files and directories contained in specified directory.

  ## Parameters

  - `dir` - list of directories to delete (default set for standard Xlsxir functionality)

  ## Example
  An example file named `test.zip` located in './test_data/test' containing a single file named `test.txt`:

      iex> path = "./test/test_data/test.zip"
      iex> file_list = ['test.txt']
      iex> Xlsxir.Unzip.extract_xml_to_file(file_list, path)
      {:ok, ['temp/test.txt']}     
      iex> Xlsxir.Unzip.delete_dir(["./temp"])
      :ok
  """
  def delete_dir(dir \\ ["temp/xl/worksheets", "temp/xl", "temp"]) do
    search_and_destroy(dir)
  end

  defp search_and_destroy([h|t]) do
    {:ok, file_list} = File.ls(h)

    case file_list do
      [] -> File.rmdir!(h)
      _  -> Enum.each(file_list, fn name -> File.rm!(h <> "/#{name}") end)
            File.rmdir!(h)
    end

    search_and_destroy(t)
  end

  defp search_and_destroy([]), do: :ok

end
