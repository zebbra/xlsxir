defmodule Xlsxir.Worksheet do
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
