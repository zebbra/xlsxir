defmodule Xlsxir.SaxParser do
  @moduledoc """
  Provides SAX (Simple API for XML) parsing functionality of the `.xlsx` file. SAX (Simple API for XML) is an event-driven 
  parsing algorithm for parsing large XML files in chunks, preventing the need to load the entire DOM into memory. Current chunk size
  is set to 10,000.
  """

  alias Xlsxir.{ParseWorksheet, ParseStyle, ParseString, Worksheet, Style, SharedString}

  @chunk 10000

  @doc """
  Parses `xl/worksheets/sheet\#{n}.xml` at index `n`, `xl/styles.xml` and `xl/sharedStrings.xml` using SAX parsing. An `Agent` 
  process is started to hold the state of data parsed. Name of `Agent` process modules are `Worksheet`, `Style` and `SharedString` 
  respectively.

  ## Parameters

  - `path` - path of XML file to be parsed in `string` format
  - `type` - file type identifier (:worksheet, :style or :string) of XML file to be parsed

  ## Example
    An example file named `test.xlsx` located in `./test/test_data` containing the following in worksheet at index `0`:
    - cell 'A1' -> "string one"
    - cell 'B1' -> "string two"
    - cell 'C1' -> integer of 10
    - cell 'D1' -> formula of `=4*5`
    - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370
    The `.xlsx` file contents have been extracted to `./test/test_data/test`. For purposes of this example, we utilize
    the `get/0` function of each `Agent` process to pull the parsed data. 

        iex> Xlsxir.SaxParser.parse("./test/test_data/test/xl/worksheets/sheet1.xml", :worksheet)
        :ok
        iex> Xlsxir.Worksheet.get
        [A1: ['s', nil, '0'], B1: ['s', nil, '1'], C1: [nil, nil, '10'], D1: [nil, nil, '20'], E1: [nil, '1', '42370']]

        iex> Xlsxir.SaxParser.parse("./test/test_data/test/xl/styles.xml", :style)
        :ok
        iex> Xlsxir.Style.get
        [nil, 'd', nil, nil, 'd', 'd']

        iex> Xlsxir.SaxParser.parse("./test/test_data/test/xl/sharedStrings.xml", :string)
        :ok
        iex> Xlsxir.String.get
        ["string one", "string two"]
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

