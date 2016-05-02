defmodule Xlsxir.Sax do
  alias Xlsxir.Worksheet

  defmodule CellState do
    defstruct cell_ref: "", data_type: "", num_style: "", value: ""
  end

  @chunk 10000

  def parse_sheet(path, type) do
    Worksheet.new
    {:ok, pid} = File.open(path, [:binary])

    index   = 0
    c_state = {pid, index, @chunk}

    :erlsom.parse_sax("",
      nil,
      case type do
        :worksheet -> &Worksheet.sax_event_handler/2
        :style     -> &Style.sax_event_handler/2
        :string    -> &String.sax_event_handler/2
        _          -> raise "Invalid file type for sax_event_handler/2"
      end,
      [{:continuation_function, &continue_file/2, c_state}])

    :ok = File.close(pid)
  end

  def continue_file(tail, {pid, offset, chunk}) do
    case :file.pread(pid, offset, chunk) do
      {:ok, data} -> {<<tail :: binary, data :: binary>>, {pid, offset + chunk, chunk}}
      :oef        -> {tail, {pid, offset, chunk}}
    end
  end

end

