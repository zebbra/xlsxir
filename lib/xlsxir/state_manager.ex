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
    Agent.start_link(fn -> [] end, name: Style)
  end

  def add_style_type(style_type) do
    Agent.update(Style, &(&1 ++ style_type)))
  end

  def get do
    Agent.get(Style, &(&1))
    |> Enum.reverse
  end

  def delete do
    Agent.stop(Style)
  end
end

defmodule Xlsxir.String do
  def new do
    Agent.start_link(fn -> [] end, name: String)
  end

  def add_shared_string(shared_string do
    Agent.update(String, &(&1 ++ shared_string)))
  end

  def get do
    Agent.get(String, &(&1))
    |> Enum.reverse
  end

  def delete do
    Agent.stop(String)
  end
end