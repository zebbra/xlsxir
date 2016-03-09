defmodule Xlsxir do

  alias Xlsxir.Unzip
  alias Xlsxir.Parse
  alias Xlsxir.Format

  def extract(path, index, option \\ :rows) do
    strings = Parse.shared_strings(path)

    path
    |> Unzip.validate_path
    |> Parse.worksheet
    |> Format.prepare_output(strings, option)
  end
end
