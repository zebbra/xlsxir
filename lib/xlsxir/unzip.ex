defmodule Xlsxir.Unzip do
  alias Xlsxir.XmlFile

  @moduledoc """
  Provides validation of accepted file types for file path,
  extracts required `.xlsx` contents to memory or files
  """
  @filetype_error "Invalid file type (expected xlsx)."
  @xml_not_found_error "Invalid File. Required XML files not found."
  @worksheet_index_error "Invalid worksheet index."

  @doc """
  Checks if given path is a valid file type and contains the requested worksheet, returning a tuple.

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format

  ## Example

         iex> path = "./test/test_data/test.xlsx"
         iex> Xlsxir.Unzip.validate_path_and_index(path, 0)
         {:ok, './test/test_data/test.xlsx'}

         iex> path = "./test/test_data/test.validfilebutnotxlsx"
         iex> Xlsxir.Unzip.validate_path_and_index(path, 0)
         {:ok, './test/test_data/test.validfilebutnotxlsx'}

         iex> path = "./test/test_data/test.xlsx"
         iex> Xlsxir.Unzip.validate_path_and_index(path, 100)
         {:error, "Invalid worksheet index."}

         iex> path = "./test/test_data/test.invalidfile"
         iex> Xlsxir.Unzip.validate_path_and_index(path, 0)
         {:error, "Invalid file type (expected xlsx)."}
  """
  def validate_path_and_index(path, index) do
    case valid_extract_request?(path, index) do
      :ok -> {:ok, path}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Checks if given path is a valid file type, returning a list of available worksheets.

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format

  ## Example

         iex> path = "./test/test_data/test.xlsx"
         iex> Xlsxir.Unzip.validate_path_all_indexes(path)
         {:ok, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]}

         iex> path = "./test/test_data/test.zip"
         iex> Xlsxir.Unzip.validate_path_all_indexes(path)
         {:ok, []}

         iex> path = "./test/test_data/test.invalidfile"
         iex> Xlsxir.Unzip.validate_path_all_indexes(path)
         {:error, "Invalid file type (expected xlsx)."}
  """

  def validate_path_all_indexes(path) do
    case list_worksheet_files(path) do
      {:ok, file_list} ->
        indexes = file_list |> Enum.with_index() |> Enum.map(fn {_, index} -> index end)
        {:ok, indexes}

      {:error, _reason} ->
        {:error, @filetype_error}
    end
  end

  def list_worksheet_files(path) do
    path = String.to_charlist(path)

    case :zip.list_dir(path) do
      {:ok, files} ->
        {:ok,
         files
         |> Enum.filter(&worksheet_file?/1)
         |> Enum.map(fn {:zip_file, path, _, _, _, _} -> path end)}

      {:error, reason} ->
        {:error, reason}
    end
  end

  defp worksheet_file?({:zip_file, filename, _, _, _, _}) do
    filename
    |> to_string()
    |> String.match?(~r/xl\/worksheets\/[Ss]heet\d+\.xml/)
  end

  defp worksheet_file?(_other), do: false

  defp valid_extract_request?(path, index) do
    case validate_path_all_indexes(path) do
      {:ok, indexes} ->
        case Enum.member?(indexes, index) do
          true -> :ok
          false -> {:error, @worksheet_index_error}
        end

      {:error, _reason} ->
        {:error, @filetype_error}
    end
  end

  @doc """
  Extracts requested list of files from a `.zip` file to memory or file system
  and returns a list of the extracted file paths.

  ## Parameters

  - `file_list` - list containing file paths to be extracted in `char_list` format
  - `path` - file path of a `.xlsx` file type in `string` format
  - `to` - `:memory` or `{:file, "destination/path"}` option

  ## Example
  An example file named `test.zip` located in './test_data/test' containing a single file named `test.txt`:

      iex> path = "./test/test_data/test.zip"
      iex> file_list = ['test.txt']
      iex> Xlsxir.Unzip.extract_xml(file_list, path, :memory)
      {:ok, [%Xlsxir.XmlFile{content: "test_successful", name: "test.txt", path: nil}]}
      iex> Xlsxir.Unzip.extract_xml(file_list, path, {:file, "temp/"})
      {:ok, [%Xlsxir.XmlFile{content: nil, name: "test.txt", path: "temp/test.txt"}]}
      iex> with {:ok, _} <- File.rm_rf("temp"), do: :ok
      :ok
  """
  def extract_xml(file_list, path, to) do
    path
    |> to_charlist
    |> extract_from_zip(file_list, to)
    |> case do
      {:error, reason} -> {:error, reason}
      {:ok, []} -> {:error, @xml_not_found_error}
      {:ok, files_list} -> {:ok, build_xml_files(files_list)}
    end
  end

  defp extract_from_zip(path, file_list, :memory),
    do: :zip.extract(path, [{:file_list, file_list}, :memory])

  defp extract_from_zip(path, file_list, {:file, dest_path}),
    do: :zip.extract(path, [{:file_list, file_list}, {:cwd, dest_path}])

  defp build_xml_files(files_list) do
    files_list
    |> Enum.map(&build_xml_file/1)
  end

  # When extracting to memory
  defp build_xml_file({name, content}) do
    %XmlFile{name: Path.basename(name), content: content}
  end

  # When extracting to temp file
  defp build_xml_file(file_path) do
    %XmlFile{name: Path.basename(file_path), path: to_string(file_path)}
  end
end
