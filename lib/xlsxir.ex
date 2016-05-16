defmodule Xlsxir do
  alias Xlsxir.{Unzip, SaxParser, Worksheet, Timer}

  @moduledoc """
  Extracts and parses data from a `.xlsx` file to an Erlang Term Storage (ETS) process and provides various functions for accessing the data.
  """

  @doc """
  Extracts worksheet data contained in the specified `.xlsx` file to an ETS process named `:worksheet` which is accessed via the `Xlsxir.Worksheet` module. Successful extraction 
  returns `:ok` with the timer argument set to false and returns a tuple of `{:ok, time}` where time is a list containing time elapsed during the extraction process 
  (i.e. `[hour, minute, second, microsecond]`) when the timer argument is set to true. 

  Cells containing formulas in the worksheet are extracted as either a `string`, `integer` or `float` depending on the resulting value of the cell. 
  Cells containing an ISO 8601 date format are extracted and converted to Erlang `:calendar.date()` format (i.e. `{year, month, day}`).

  ## Parameters
  - `path` - file path of a `.xlsx` file type in `string` format
  - `index` - index of worksheet from within the Excel workbook to be parsed (zero-based index)
  - `timer` - boolean flag that tracts extraction process time and returns it when set to `true`. Defalut value is `false`.

  ## Example
  Extract first worksheet in an example file named `test.xlsx` located in `./test/test_data`:

        iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
        :ok
        iex> Xlsxir.Worksheet.alive?
        true
        iex> Xlsxir.close
        :ok
  """
  def extract(path, index, timer \\ false) do
    if timer, do: Timer.start

    {:ok, file}       = Unzip.validate_path(path)
    {:ok, file_paths} = Unzip.xml_file_list(index)
                        |> Unzip.extract_xml_to_file(file)

    Enum.each(file_paths, fn file -> 
      case file do
        'temp/xl/sharedStrings.xml' -> SaxParser.parse(to_string(file), :string)
        'temp/xl/styles.xml'        -> SaxParser.parse(to_string(file), :style)
        _                           -> nil
      end
    end)

    SaxParser.parse("temp/xl/worksheets/sheet#{index + 1}.xml", :worksheet)
    Unzip.delete_dir

    if timer, do: {:ok, Timer.stop}, else: :ok
  end

  @doc """
  Accesses `:worksheet` ETS process and returns data formatted as a list of row value lists.

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
          :ok
          iex> Xlsxir.get_list
          [["string one", "string two", 10, 20, {2016, 1, 1}]]
          iex> Xlsxir.close
          :ok
  """
  def get_list do
    :ets.match(:worksheet, {:"$1", :"$2"})
    |> Enum.sort
    |> Enum.map(fn obj -> Enum.at(obj, 1)
                          |> Enum.map(fn row -> Enum.at(row, 1) end)
                        end)
  end

  @doc """
  Accesses `:worksheet` ETS process and returns data formatted as a map of cell references and values.

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
          :ok
          iex> Xlsxir.get_map
          %{ "A1" => "string one", "B1" => "string two", "C1" => 10, "D1" => 20, "E1" => {2016,1,1}}
          iex> Xlsxir.close
          :ok
  """
  def get_map do
    :ets.match(:worksheet, {:"$1", :"$2"})
    |> Enum.sort
    |> Enum.reduce(%{}, fn match_obj, acc -> 
         Enum.at(match_obj, 1)
         |> Enum.reduce(%{}, fn [k, v], acc2 -> Map.put(acc2, k, v) end)
         |> Enum.into(acc)
       end)
  end

  @doc """
  Accesses `:worksheet` ETS process and returns value of specified cell.

  ## Parameters
  - `cell_ref` - Reference name of cell to be returned in `string` format (i.e. `"A1"`)

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
          :ok
          iex> Xlsxir.get_cell("A1")
          "string one"
          iex> Xlsxir.close
          :ok
  """
  def get_cell(cell_ref) do
    [[row_num]] = ~r/\d+/ |> Regex.scan(cell_ref)
    [[row]]     = :ets.match(:worksheet, {row_num, :"$1"})

    row
    |> Enum.filter(fn cell -> Enum.at(cell, 0) == cell_ref end) 
    |> List.first
    |> Enum.at(1)
  end

  @doc """
  Accesses `:worksheet` ETS process and returns values of specified row in a `list`.

  ## Parameters
  - `row` - Reference name of row to be returned in `integer` format (i.e. `1`)

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
          :ok
          iex> Xlsxir.get_row(1)
          ["string one", "string two", 10, 20, {2016, 1, 1}]
          iex> Xlsxir.close
          :ok
  """
  def get_row(row) do
    [[row]] = :ets.match(:worksheet, {to_string(row), :"$1"})

    row
    |> Enum.map(fn cell -> Enum.at(cell, 1) end)
  end

  @doc """
  Accesses `:worksheet` ETS process and returns values of specified column in a `list`.

  ## Parameters
  - `col` - Reference name of column to be returned in `string` format (i.e. `"A"`)

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
          :ok
          iex> Xlsxir.get_col("A")
          ["string one"]
          iex> Xlsxir.close
          :ok
  """
  def get_col(col) do
    Enum.map(get_map, fn {k,v} -> if cell_ltrs(k) == col, do: v end)
    |> Enum.reject(fn x -> x == nil end)
  end

  @doc """
  Deletes ETS process `:worksheet` and returns `:ok` if successful.

  ## Example
  Extract first worksheet in an example file named `test.xlsx` located in `./test/test_data`:

      iex> Xlsxir.extract("./test/test_data/test.xlsx", 0)
      :ok
      iex> Xlsxir.close
      :ok
  """
  def close do
    Worksheet.delete
    |> case do
      false -> raise "Unable to close worksheet"
      true -> :ok
    end
  end

  defp cell_ltrs(cell) do
    ~r/[a-z]+/i 
    |> Regex.scan(String.upcase(cell))
    |> List.flatten
    |> List.first
  end

end
