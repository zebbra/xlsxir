defmodule Xlsxir.SaxParser do
  @moduledoc """
  
  """

  alias Xlsxir.{ParseWorksheet, ParseStyle, ParseString, Worksheet, Style, SharedString}

  @chunk 10000

  @doc """
  
  """
  def parse(path, type) do
    case type do
      :worksheet -> Worksheet.new
      :style     -> Style.new
      :string    -> SharedString.new
    end

    {:ok, pid} = File.open(path, [:binary])

    index   = 0
    c_state = {pid, index, @chunk}

    :erlsom.parse_sax("",
      nil,
      case type do
        :worksheet -> &ParseWorksheet.sax_event_handler/2
        :style     -> &ParseStyle.sax_event_handler/2
        :string    -> &ParseString.sax_event_handler/2
        _          -> raise "Invalid file type for sax_event_handler/2"
      end,
      [{:continuation_function, &continue_file/2, c_state}])

    :ok = File.close(pid)
  end

  defp continue_file(tail, {pid, offset, chunk}) do
    case :file.pread(pid, offset, chunk) do
      {:ok, data} -> {<<tail :: binary, data :: binary>>, {pid, offset + chunk, chunk}}
      :oef        -> {tail, {pid, offset, chunk}}
    end
  end

end

