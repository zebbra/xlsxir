defmodule Xlsxir.ConvertDate do
  @moduledoc """
  Converts an ISO 8601 date format serial number, in `char_list` format, to a date formatted in 
  Erlang `:calendar.date()` type format (i.e. `{year, month, day}`).
  """

  @doc """
  Receives an ISO 8601 date format serial number and returns a date formatted in Erlang `:calendar.date()`
  type format.

  ## Parameters

  - `serial` - ISO 8601 date format serial in `char_list` format (i.e. 4/30/75 as '27514')

  ## Example

      iex> Xlsxir.ConvertDate.from_serial('27514')
      {1975, 4, 30}
  """
  def from_serial(serial) do
    f_serial = serial
               |> convert_char_number
               |> is_float
               |> case do
                    false -> List.to_integer(serial)
                    true  -> List.to_float(serial)
                             |> Float.floor
                             |> round
                  end
               
    f_serial
    |> process_serial_int
    |> determine_year
    |> determine_month_and_day
  end

  defp process_serial_int(serial_int) do
    year = serial_int 
           |> Kernel./(365) 
           |> Float.floor
           |> Kernel.+(1900) 
           |> round

    serial_int = if serial_int >= 60 && serial_int <= 364, do: serial_int - 1, else: serial_int

    days  = serial_int
            |> rem(365)
            |> Kernel.-(
                        serial_int
                        |> Kernel./(365)
                        |> Float.floor
                        |> Kernel./(4)
                        |> Float.ceil
                       )

    {year, days}
  end

  defp determine_year({year, days}) do
    if days <= 0, do: {year - 1, days}, else: {year, days}  
  end

  defp determine_month_and_day({year, days}) do
    l = if :calendar.is_leap_year(year), do: 1, else: 0


    {month, day} = if days <= 0 do
                     {12, 31 + days}
                   else
                     process_days(days, l)
                   end

    {year, month, round(day)}
  end

  defp process_days(days, l) do
    cond do
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
      days <= 365 + l -> {12, days - 334 - l}
      true            -> raise "Invalid Excel serial date."
    end
  end

  @doc """
  Converts extracted number in `char_list` format to either `integer` or `float`.
  """
  def convert_char_number(number) do
    number
    |> List.to_string
    |> String.match?(~r/[.]/)
    |> case do
         false -> List.to_integer(number)
         true  -> List.to_float(number)
       end
  end
end