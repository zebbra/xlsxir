# Xlsxir

Xlsxir is an Elixir library that parses `.xlsx` files using Simple API for XML (SAX) parsing via the [Erlsom](https://github.com/willemdj/erlsom) Erlang library, extracts the data to an Erlang Term Storage (ETS) process and provides various functions for accessing the data. Xlsxir supports ISO 8601 date formats and large files. Testing has been limited to various documents I have created or have access to and any issues submitted through GitHub, though I have succesfully parsed a worksheet containing 100 rows and 514K columns. Please submit any issues found and they will be addressed ASAP.  

## Installation

You can add Xlsxir as a dependancy to your Elixir project via the Hex package manager by adding the following to your `mix.exs` file: 

```elixir
def deps do
  [ {:xlsxir, "~> 1.0.0"} ]
end
```

Or, you can directly reference the GitHub repo:

```elixir
def deps do
  [ {:xlsxir, github: "kennellroxco/xlsxir"} ]
end
```

## Basic Usage

Xlsxir parses a `.xlsx` file located at a given `path` and extracts the data to an ETS process via the `Xlsxir.extract/3` function:

```elixir
Xlsxir.extract(path, index, timer \\ false)
```

Argument descriptions:
- `path` the path of the file to be parsed in `string` format
- `index` is the position of the worksheet you wish to parse (zero-based index)
- `timer` is a boolean flag that controls an extraction timer that returns time elapsed when set to `true`. Defalut value is `false`.

Upon successful completion, the extraction process returns `:ok` with `timer` set to `false`, or `{:ok, time_elapsed}` with `timer` set to `true`.

The extracted worksheet data can be accessed using any of the following functions:
- `Xlsxir.get_list/0` - Returns entire worksheet data in the form of a list of row lists (i.e. `[[row 1 values], [row 2 values], ...]`)
- `Xlsxir.get_map/0` - Returns entire worksheet data in the form of a map of cell names and values (i.e. `%{"A1" => value, "A2" => value, ...}`)
- `Xlsxir.get_cell/1` - Returns value of specified cell (i.e. `"A1"` returns value contained in cell A1)
- `Xlsxir.get_row/1` - Returns values of specified row (i.e. `1` returns the first row of data)
- `Xlsxir.get_col/1` - Returns values of specified column (i.e `"A"` returns the first column of data)

Once the table data is no longer needed, run `Xlsxir.close` to delete the ETS process and free memory.

Refer to [Xlsxir documentation](https://hexdocs.pm/xlsxir/index.html) for more detailed examples. 

## Considerations

Cell references are formatted as a string (i.e. "A1"). Strings will be returned as type `string`, resulting values for functions from within the worksheet are returned as type `string`, `integer` or `float` depending on the type of the resulting value, data formatted as a number in the worksheet will be returned as type `integer` or `float`, and ISO 8601 date formatted values will be returned in Erlang `:calendar.date()` type format (i.e. `{year, month, day}`). Xlsxir does not currently support dates prior to 1/1/1900.

## Planned Development

- Additional performance improvement for larger files
- Adding time support for dates (i.e. {{YYYY, MM, DD}, {h, m, s}})
- Export functionality to .xlsx file type with formatting options
- Implement Elixir 1.3 calendar datatypes support

## Contributing

Contributions are encouraged. Feel free to fork the repo, add your code along with appropriate tests and documentation (ensuring all existing tests continue to pass) and submit a pull request. 

## Bug Reporting

Please report any bugs or request future enhancements via the [Issues](https://github.com/kennellroxco/xlsxir/issues) page. 
