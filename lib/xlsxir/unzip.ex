defmodule Xlsxir.Unzip do

  @spec validate_path(String.t) :: {:ok | :error, String.t}
  def validate_path(path) do
    path
    |> String.downcase
    |> String.split(".", trim: true)
    |> List.last
    |> case do
      "xlsx" -> {:ok, path}
      _      -> {:error, "Invalid path. Currently only .xlsx is supported."}
    end  
  end

  @spec extract_xml(String.t, [char]) :: {:ok | :error, String.t}
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
          {:error, cause} -> {:error, cause}
          {:ok, {_, file_content}} ->
            case :zip.zip_close(zip_directory) do
              {:error, cause} -> {:error, cause}
              :ok             -> {:ok, file_content}
            end
        end
    end
  end

end