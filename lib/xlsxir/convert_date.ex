defmodule Xlsxir.ConvertDate do

  @moduledoc """
  Converts an Excel date serial number, in `char_list` format, to a date formatted in 
  Erlang `:calendar.date()` type format (i.e. `{year, month, day}`).
  """

  @doc """
  Receives an Excel date serial number and returns a date formatted in Erlang `:calendar.date()`
  type format.

  ## Parameters

  - `serial` - Excel date serial in `char_list` format (i.e. 4/30/75 as '27514')

  ## Example

      iex> Xlsxir.ConvertDate.from_excel('27514')
      {1975, 4, 30}
  """
  def from_excel(serial) do
    serial
    |> List.to_integer
    |> calc_year
    |> calc_month_and_day
  end

  defp calc_year(serial) do
    years = Float.floor(serial/365) 
    leap_years = Float.ceil(years/4)
    days = rem(serial, 365) - leap_years
    year = round(1900 + years)

    {year, days}
  end

  defp calc_month_and_day({year, days}) do
    l = if :calendar.is_leap_year(year), do: 1, else: 0

    {month, day} = cond do
                     days <= 31      -> {1, days}
                     days <= 59 + l  -> {2, days - 31}
                     days <= 90 + l  -> {3, days - 59 - l}
                     days <= 120 + l -> {4, days - 90 - l}
                     days <= 151 + l -> {5, days - 120 - l}
                     days <= 181 + l -> {6, days - 151 - l}
                     days <= 212 + l -> {7, days - 181 - l}
                     days <= 243 + l -> {8, days - 212 - l}
                     days <= 273 + l -> {9, days - 243 - l}
                     days <= 304 + l -> {10, days - 273 - l}
                     days <= 334 + l -> {11, days - 304 - l}
                     true           -> {12, days - 334 - l}
                   end

    {year, month, round(day)}
  end
end