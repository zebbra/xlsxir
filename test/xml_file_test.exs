defmodule XmlFileTest do
  use ExUnit.Case

  test "open memory XmlFile" do
    assert {:ok, _file_pid} = Xlsxir.XmlFile.open(%Xlsxir.XmlFile{content: File.read!("./test/test_data/test/xl/styles.xml")})
  end

  test "open filepath XmlFile" do
    assert {:ok, _file_pid} = Xlsxir.XmlFile.open(%Xlsxir.XmlFile{path: "./test/test_data/test/xl/styles.xml"})
  end
end
