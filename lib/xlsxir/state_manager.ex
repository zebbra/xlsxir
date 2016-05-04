defmodule Xlsxir.Worksheet do
  @moduledoc """
  An `Agent` process named `Sheet` which holds state for data parsed from `sheet\#{n}.xml` at index `n`. Provides 
  functions to create the process, add and retreive data, and ultimately kill the process. 
  """
  def new do
    Agent.start_link(fn -> [] end, name: Sheet)
  end

  def add_cell(key, value) do
    Agent.update(Sheet, &(Keyword.put(&1, key, value)))
  end

  def get do
    Agent.get(Sheet, &(&1))
    |> Enum.reverse
  end

  def delete do
    Agent.stop(Sheet)
  end
end

defmodule Xlsxir.SharedString do
  @moduledoc """
  An `Agent` process named `SharedString` which holds state for data parsed from `sharedStrings.xml`. Provides 
  functions to create the process, add and retreive data, and ultimately kill the process.
  """
  def new do
    Agent.start_link(fn -> [] end, name: SharedStrings)
  end

  def add_shared_string(shared_string) do
    Agent.update(SharedStrings, &(Enum.into([shared_string], &1)))
  end

  def get do
    Agent.get(SharedStrings, &(&1))
  end

  def delete do
    Agent.stop(SharedStrings)
  end
end

defmodule Xlsxir.Style do
  @moduledoc """
  An `Agent` process named `Styles` which holds state for data parsed from `styles.xml`. Provides 
  functions to create the process, add and retreive data, and ultimately kill the process. Also includes
  a temporary `Agent` process named `NumFmtIds` which is utilized during the parsing of the `styles.xml` file
  to temporarily hold state of each `NumFmtId` contained within the file. 
  """
  def new do
    Agent.start_link(fn -> [] end, name: NumFmtIds)
    Agent.start_link(fn -> [] end, name: Styles)
  end

  # functions for `NumFmtIds`
  def add_id(num_fmt_id) do
    unless Enum.member?(get_id, num_fmt_id) do 
      Agent.update(NumFmtIds, &(Enum.into([num_fmt_id], &1)))
    end
  end

  def get_id do
    Agent.get(NumFmtIds, &(&1))
  end

  def delete_id do
    Agent.stop(NumFmtIds)
  end

  # functions for `Styles`
  def add(style) do
    Agent.update(Styles, &(Enum.into([style], &1)))
  end

  def get do
    Agent.get(Styles, &(&1))
  end

  def delete do
    Agent.stop(Styles)
  end

  def alive? do
    case Process.whereis(Styles) do
      nil -> false
      pid -> Process.alive?(pid)
    end
  end
end

