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

defmodule Xlsxir.Style do
  def new do
    Agent.start_link(fn -> [] end, name: NumFmtId)
    Agent.start_link(fn -> [] end, name: Style)
  end

  def add_id(num_fmt_id) do
    unless Enum.member?(get_id, num_fmt_id), do: Agent.update(NumFmtId, &(Enum.into([num_fmt_id], &1)))
  end

  def add_style(style) do
    Agent.update(Style, &(Enum.into([style], &1)))
  end

  def get_id do
    Agent.get(NumFmtId, &(&1))
  end

  def get_style do
    Agent.get(Style, &(&1))
  end

  def delete_id do
    Agent.stop(NumFmtId)
  end

  def delete_style do
    Agent.stop(Style)
  end
end

defmodule Xlsxir.SharedString do
  def new do
    Agent.start_link(fn -> [] end, name: String)
  end

  def add_shared_string(shared_string) do
    Agent.update(String, &(Enum.into([shared_string], &1)))
  end

  def get do
    Agent.get(String, &(&1))
  end

  def delete do
    Agent.stop(String)
  end
end