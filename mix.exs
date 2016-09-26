defmodule Xlsxir.Mixfile do
  use Mix.Project

  def project do
    [
     app: :xlsxir,
     version: "1.3.5",
     name: "Xlsxir",
     source_url: "https://github.com/kennellroxco/xlsxir",
     elixir: "~> 1.2",
     build_embedded: Mix.env == :prod,
     start_permanent: Mix.env == :prod,
     description: description,
     package: package,
     deps: deps,
     docs: [main: "overview", extras: ["CHANGELOG.md", "NUMBER_STYLES.md", "OVERVIEW.md"]] 
    ]
  end

  def application do
    [applications: [:logger]]
  end

  defp deps do
    [ 
      { :ex_doc, github: "elixir-lang/ex_doc", only: :dev },
      { :earmark, github: "pragdave/earmark", only: :dev },
      { :erlsom, "~> 1.4" }
    ]
  end

  defp description do
    """
    Xlsx file parser. Supports large files, multiple worksheets and ISO 8601 date formats. Data is extracted to an Erlang Term Storage (ETS) table and is accessed through various functions. Tested with Excel and LibreOffice.
    """
  end

  defp package do
    [
      maintainers: ["Jason Kennell", "Bryan Weatherly"],
      licenses: ["MIT License"],
      links: %{
                "Github" => "https://github.com/kennellroxco/xlsxir",
                "Change Log" => "https://hexdocs.pm/xlsxir/changelog.html"
               }
    ]
  end
  
end
