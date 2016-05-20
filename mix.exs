defmodule Xlsxir.Mixfile do
  use Mix.Project

  def project do
    [
     app: :xlsxir,
     version: "1.2.0",
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

  # Configuration for the OTP application
  #
  # Type "mix help compile.app" for more information
  def application do
    [applications: [:logger]]
  end

  # Dependencies can be Hex packages:
  #
  #   {:mydep, "~> 0.3.0"}
  #
  # Or git/path repositories:
  #
  #   {:mydep, git: "https://github.com/elixir-lang/mydep.git", tag: "0.1.0"}
  #
  # Type "mix help deps" for more examples and options
  defp deps do
    [ 
      { :ex_doc,    "~> 0.11.4" },
      { :earmark,   "~> 0.2.1"  },
      { :erlsom,    "~> 1.4"    }
    ]
  end

  defp description do
    """
    Xlsx file parser. Supports large files and ISO 8601 date formats. Data is extracted to an Erlang Term Storage (ETS) table and is accessed through various functions.
    """
  end

  defp package do
    [
      maintainers: ["Jason Kennell", "Bryan Weatherly"],
      licenses: ["MIT License"],
      links: %{"Github" => "https://github.com/kennellroxco/xlsxir"}
    ]
  end
  
end
