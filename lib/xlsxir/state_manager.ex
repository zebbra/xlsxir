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
    row = :ets.lookup(:worksheet, index)

    if row == [] do
      row
    else
      row
      |> List.first
      |> Tuple.to_list
      |> Enum.at(1)
    end
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
    Agent.update(NumFmtIds, &(Enum.into([num_fmt_id], &1)))
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
    Agent.start_link(fn -> [] end, name: Time)
    {_, s, ms} = :erlang.timestamp
    Agent.update(Time, &(Enum.into([s, ms],&1)))
  end

  def stop do
    {_, s, ms} = :erlang.timestamp

    seconds = s |> Kernel.-(Agent.get(Time, &(&1)) |> Enum.at(0))
    micro   = ms |> Kernel.-(Agent.get(Time, &(&1)) |> Enum.at(1))

    hms = [
            seconds/3600 |> Float.floor |> round, 
            rem(seconds, 3600)/60 |> Float.floor |> round, 
            rem(seconds, 60)
          ]

    case hms do
      [0, 0, 0] -> "#{micro}ms"
      [0, 0, s] -> "#{s}s #{micro}ms"
      [0, m, s] -> "#{m}m #{s}s #{micro}ms"
      [h, m, s] -> "#{h}h #{m}m #{s}s #{micro}ms"
    end
  end
end

