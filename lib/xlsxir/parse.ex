defmodule Xlsxir.Parse do
  import Xlsxir.Unzip, only: [extract_xml: 2]
  import SweetXml

  @moduledoc """
  Receives Excel xml data via the `extract_xml` function of the `Unzip` module and parses it.
  """

  @doc """
  Receives Excel string data in xml format, parses it and returns the strings in the form of a list.

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format

  ## Example
    An example file named `test.xlsx` located in `./test/test_data` containing the following:
    - cell 'A1' -> "string one"
    - cell 'B1' -> "string two"
    - cell 'C1' -> integer of 10
    - cell 'D1' -> formula of `=4*5`
    - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> Xlsxir.Parse.shared_strings("./test/test_data/test.xlsx")
          ["string one", "string two"]
  """
  def shared_strings(path) do
    {:ok, strings} = extract_xml(path, 'xl/sharedStrings.xml')
    strings
    |> xpath(~x"//t"l)
    |> Enum.map(fn string -> case string do
          {:xmlElement,_,_,_,_,_,_,_,[{_,_,_,_,str,_}],_,_,_} -> to_string(str)
          {:xmlElement,_,_,_,_,_,_,_,_,_,_,_}                 -> join_string_fragments(string)
          _                                                   -> raise "sharedStrings.xml parse error"
        end 
      end)
  end

  defp join_string_fragments(xml) do
    Tuple.to_list(xml)
    |> Enum.at(8)
    |> Enum.reduce("", fn(x, acc) -> {_,_,_,_,str,_} = x
        acc <> to_string(str)
      end)
  end

  @doc """
  Receives Excel style data in xml format, parses it and returns the `numFmtId` attributes in a list which is then processed
  into a list of style types (`'d'` for date type, `'nil'` for standard number type).

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format

  ## Example
    The example file named `test.xlsx` located in `./test/test_data` contains the following `numFmtId` attributes:
    - 1 at index 0 which is the standard format for Excel numbers
    - 14 at index 1 which is the Excel date format of `mm-dd-yy`

          iex> Xlsxir.Parse.num_style("./test/test_data/test.xlsx")
          [nil, 'd']
  """
  def num_style(path) do
    {:ok, styles} = extract_xml(path, 'xl/styles.xml')
    custom = custom_style(styles)

    styles
    |> xpath(~x"//cellXfs/xf/@numFmtId"l)
    |> Enum.map(fn style_type -> 
        case List.to_integer(style_type) do
          i when i in 0..4           -> nil
          i when i in 9..13          -> nil
          i when i in 14..22         -> 'd'
          i when i in 37..40         -> nil
          i when i in 45..47         -> 'd'
          i when i in 48..49         -> nil
          i when i in 59..62         -> nil
          i when i in 67..70         -> nil
          27 || 30 || 36 || 50 || 57 -> 'd'
          44                         -> nil
          i when i > 4               -> custom[style_type] 
          _                          -> raise "Unsupported style type: #{style_type}."
        end                          
      end)
  end

  defp custom_style(xml) do
    custom = xml
             |> xpath(~x"//numFmt"l)
             |> Enum.reduce(%{}, fn fmt, acc -> 
                 {_,_,_,_,_,_,_,[{_,:numFmtId,_,_,_,_,_,_,id,_},{_,:formatCode,_,_,_,_,_,_,code,_}],_,_,_,_} = fmt
                 Map.put_new(acc, id, code)
               end)

    custom
    |> Enum.reduce(%{}, fn {k, v}, acc -> 
         cond do
           String.match?(to_string(v), ~r/\bred\b/i) -> Map.put_new(acc, k, nil)
           String.match?(to_string(v), ~r/[dhmsy]/)  -> Map.put_new(acc, k, 'd')
           true                                      -> Map.put_new(acc, k, nil)
         end
      end)
  end

  @doc """
  Receives the xlsx worksheet at position `index` in xml format, parses the data and returns
  required elements in the form of a `keyword list`:

      [[row_1_cell_1: ['attribute', 'value'], ...], [row_2_cell_1: ['attribute', 'value'], ...], ...]

  ## Parameters

  - `path` - file path of a `.xlsx` file type in `string` format
  - `index` - index of worksheet from within the Excel workbook to be parsed, starting with `0`

  ## Example
    An example file named `test.xlsx` located in `./test/test_data` containing the following in worksheet at index `0`:
    - cell 'A1' -> "string one"
    - cell 'B1' -> "string two"
    - cell 'C1' -> integer of 10
    - cell 'D1' -> formula of `=4*5`
    - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> Xlsxir.Parse.worksheet("./test/test_data/test.xlsx", 0, [nil,'d'])
          [[A1: ['s', nil, '0'], B1: ['s', nil, '1'], C1: [nil, nil, '10'], D1: [nil, nil, '20'], E1: [nil, 'd', '42370']]]
  """
  def worksheet(path, index, styles) do
    {:ok, sheet} = extract_xml(path, 'xl/worksheets/sheet#{index + 1}.xml')

    sheet
    |> xpath(~x"//worksheet/sheetData/row/c"l)
    |> Stream.map(&(process_column(&1, styles)))
    |> Enum.chunk_by(fn cell -> Keyword.keys([cell])
                                |> List.first
                                |> Atom.to_string
                                |> regx_scan
                              end)
  end

  defp process_column({:xmlElement,:c,:c,_,_,_,_,xml_attr,xml_elem,_,_,_}, styles) do
    {cell_ref, num_style, data_type} = extract_attribute(xml_attr, styles)
    {List.to_atom(cell_ref), [data_type, num_style, extract_value(xml_elem)]}
  end

  defp extract_attribute(xml_attr, styles) do
    a = Enum.map(xml_attr, fn(attr) -> 
      case attr do
        {:xmlAttribute,:r,_,_,_,_,_,_,ref,_}   -> {:r, ref}
        {:xmlAttribute,:s,_,_,_,_,_,_,style,_} -> {:s, Enum.at(styles, List.to_integer(style))}
        {:xmlAttribute,:t,_,_,_,_,_,_,type,_}  -> {:t, type}
        _                                      -> raise "Unknown cell attribute"
      end
    end)

    {cell_ref, num_style, data_type} = case Keyword.keys(a) do
                                        [:r]         -> {a[:r],   nil,   nil}
                                        [:r, :s]     -> {a[:r], a[:s],   nil}
                                        [:r, :t]     -> {a[:r],   nil, a[:t]}
                                        [:r, :s, :t] -> {a[:r], a[:s], a[:t]}
                                        _            -> raise "Invalid attributes: #{a}"
                                       end
    {cell_ref, num_style, data_type}
  end

  defp extract_value(xml_elem) do
    case xml_elem do
      [{:xmlElement,_,_,_,_,_,_,_,[{_,_,_,_,val,_}],_,_,_}]         -> val
      [_,{:xmlElement,_,_,_,_,_,_,_,[{_,_,_,_,funct_val,_}],_,_,_}] -> funct_val
      []                                                            -> nil
      _                                                             -> raise "Invalid xmlElement."
    end
  end

  defp regx_scan(cell) do
    ~r/[0-9]/
    |> Regex.scan(cell)
    |> List.to_string
  end

end
