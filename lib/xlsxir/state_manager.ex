defmodule Xlsxir.Worksheet do
  @moduledoc """
  An Erlang Term Storage (ETS) process named `:worksheet` which holds state for data parsed from `sheet\#{n}.xml` at index `n`. Provides 
  functions to create the process, add and retreive data, and ultimately kill the process. 
  """

  def new do
    :ets.new(:worksheet, [:set, :protected, :named_table])
  end

  def add_row(row, index) do
    :worksheet |> :ets.insert({index, row})
  end

  def get_at(index) do
    :ets.lookup(:worksheet, index)
    |> List.first
    |> Tuple.to_list
    |> Enum.at(1)
  end

  def delete do
    :ets.delete(:worksheet)
  end
end

defmodule Xlsxir.SharedString do
  @moduledoc """
  An Erlang Term Storage (ETS) process named `:sharedstrings` which holds state for data parsed from `sharedStrings.xml`. Provides 
  functions to create the process, add and retreive data, and ultimately kill the process.
  """

  def new do
    :ets.new(:sharedstrings, [:set, :protected, :named_table])
  end

  def add_shared_string(shared_string, index) do
    :sharedstrings |> :ets.insert({index, shared_string})
  end

  def get_at(index) do
    :ets.lookup(:sharedstrings, index)
    |> List.first
    |> Tuple.to_list
    |> Enum.at(1)
  end

  def delete do
    :ets.delete(:sharedstrings)
  end

  def alive? do
    Enum.member?(:ets.all, :sharedstrings)
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
    :ets.new(:styles, [:set, :protected, :named_table])
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

  # functions for `:styles`
  def add_style(style, index) do
    :styles |> :ets.insert({index, style})
  end

  def get_at(index) do
    :ets.lookup(:styles, index)
    |> List.first
    |> Tuple.to_list
    |> Enum.at(1)
  end

  def delete do
    :ets.delete(:styles)
  end

  def alive? do
    Enum.member?(:ets.all, :styles)
  end
end

defmodule Xlsxir.Index do
  @moduledoc """
  An `Agent` process named `Index` which holds state of an index. Provides functions to create the process, 
  increment the index by 1, and ultimately kill the process.
  """

  def new do
    Agent.start_link(fn -> 0 end, name: Index)
  end

  def increment_1 do
    Agent.update(Index, &(&1 + 1))
  end

  def get do
    Agent.get(Index, &(&1))
  end

  def delete do
    Agent.stop(Index)
  end
end

defmodule Xlsxir.Timer do
  @moduledoc """
  An `Agent` process named `Time` which holds state for time elapsed since execution. Provides functions to create the process, 
  start the timer, stop the timer, reset, restart, and ultimately kill the process.
  """

  def start do
    Agent.start_link(fn -> 0 end, name: Time)
    {_, s, _} = :erlang.now
    Agent.update(Time, &(&1 + s))
  end

  def restart do
    reset
    {_, s, _} = :erlang.now
    Agent.update(Time, &(&1 + s))
  end

  def stop do
    {_, s, _} = :erlang.now

    s
    |> Kernel.-(Agent.get(Time, &(&1)))
    |> convert_seconds
  end

  def reset do
    Agent.update(Time, &(&1 = 0))
  end

  def delete do
    Agent.stop(Time)
  end

  defp convert_seconds(seconds) do
    [h, m, s] = [
                  seconds/3600 |> Float.floor |> round, 
                  rem(seconds, 3600)/60 |> Float.floor |> round, 
                  rem(seconds, 60)
                ]

    "#{h}h #{m}m #{s}s"            
  end
end

