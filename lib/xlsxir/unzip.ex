defmodule Xlsxir.Unzip do

  @moduledoc """
  Provides validation of accepted file types for file path, extracts required `.xlsx` contents to `./temp` and ultimately deletes the `./temp` directory and its contents.
  """

  @doc """
  Checks if given path is a valid file type and contains the requested worksheet, returning a tuple.

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format

  ## Example

         iex> path = "./test/test_data/test.xlsx"
         iex> Xlsxir.Unzip.validate_path(path, 0)
         {:ok, './test/test_data/test.xlsx'}
        
         iex> path = "./test/test_data/test.validfilebutnotxlsx"
         iex> Xlsxir.Unzip.validate_path(path, 0)
         {:ok, './test/test_data/test.validfilebutnotxlsx'}

         iex> path = "./test/test_data/test.xlsx"
         iex> Xlsxir.Unzip.validate_path(path, 100)
         {:error, "Invalid path. Currently only .xlsx file types are supported."}
  """
  def validate_path(path, index) do
    path = String.to_char_list(path)
    {:ok, file_list} = :zip.list_dir(path)

    if search_file_list(file_list, index) do
      {:ok, path}
    else
      {:error, "Invalid path. Currently only .xlsx file types are supported."}
    end
  end

  defp search_file_list(file_list, index) do
    sheet = 'xl/worksheets/sheet#{index + 1}.xml'

    file_list
    |> Enum.any?(fn file -> 
         case file do
           {:zip_file, ^sheet, _, _, _, _} -> true
           _                               -> false
         end
       end)
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
      iex> Xlsxir.Unzip.delete_dir("./temp")
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
      iex> Xlsxir.Unzip.delete_dir("./temp")
      :ok
  """
  def delete_dir(dir \\ "./temp") do
    File.rm_rf(dir)
    :ok
  end

end
