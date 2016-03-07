defmodule Xlsxir.Parse do
  import SweetXml
  
  def shared_strings_with_index(path) do
    {:ok, strings} = extract_xml(path, 'xl/sharedStrings.xml')
    strings
    |> xpath(~x"//t/text()"sl)
    |> Enum.with_index
  end

  def worksheet(path, index) do
    {:ok, worksheet} = extract_xml(path, 'xl/worksheets/sheet#{index + 1}.xml')
    worksheet
    |> xmap(
        sheet: [
          ~x"//row"l,
          row: ~x"./@r",
          columns: [
            ~x"./c"l,
            column:   ~x"./@r",
            type:     ~x"./@t",
            function: ~x"./f/text()"o,
            value:    ~x"./v/text()"
            ] 
          ]
        )
  end

  defp extract_xml(path, inner_path) do
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