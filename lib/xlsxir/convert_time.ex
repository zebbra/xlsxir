defmodule Xlsxir.ConvertTime do
  @moduledoc """
  Converts a time formatted as a decimal number that represents the fraction
  of the day in `char_list` form, to a tiem formatted in Erlang `:erlang.time()`
  format (i.e. `{hour, minute, second}`).
  """

  @doc """
  Given a charlist in the form of a fraction of the day, return the time
  formatted as `:erlang.time()`.

  Note that the `float` representation of the input must be a number between
  `0` (inclusive) and `1` (exclusive).

  ## Parameters

  - `charlist` - Character list in the form of a fractional number (i.e. `0.25`)

  ## Example

      iex> Xlsxir.ConvertTime.from_charlist('0.25')
      {6, 0, 0}
  """
  def from_charlist('0'), do: {0, 0, 0}
  def from_charlist(charlist) do
    charlist
    |> List.to_float
    |> from_float
  end

  defp from_float(time)
  when is_float(time) and time >= 0 and time < 1.0 do
    {hours, min_fraction} = split_float(time * 24)
    {minutes, sec_fraction} = split_float(min_fraction * 60)
    {seconds, _} = split_float(sec_fraction * 60)

    {hours, minutes, seconds}
  end

  defp split_float(x) do
    a = round(Float.floor(x))
    {a, x - a}
  end
end
