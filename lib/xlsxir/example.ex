defmodule Xlsxir.Example do
  
  def sheet do
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
      <dimension ref="A1:D5"/>
      <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0">
          <selection activeCell="D6" sqref="D6"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="16" x14ac:dyDescent="0.2"/>
      <cols>
        <col min="4" max="4" width="12.1640625" bestFit="1" customWidth="1"/>
      </cols>
      <sheetData>
        <row r="1" spans="1:4" x14ac:dyDescent="0.2">
          <c r="A1" t="s">
            <v>0</v>
          </c>
          <c r="B1" t="s">
            <v>1</v>
          </c>
          <c r="C1" t="s">
            <v>2</v>
          </c>
          <c r="D1" t="s">
            <v>3</v>
          </c>
        </row>
        <row r="2" spans="1:4" x14ac:dyDescent="0.2">
          <c r="A2" t="s">
            <v>4</v>
          </c>
          <c r="B2" t="s">
            <v>5</v>
          </c>
          <c r="C2" t="s">
            <v>6</v>
          </c>
          <c r="D2" t="s">
            <v>7</v>
          </c>
        </row>
        <row r="3" spans="1:4" x14ac:dyDescent="0.2">
          <c r="A3">
            <v>1</v>
          </c>
          <c r="B3">
            <v>345</v>
          </c>
          <c r="C3" s="1">
            <v>42377</v>
          </c>
          <c r="D3">
            <f>4*5</f>
            <v>20</v>
          </c>
        </row>
        <row r="4" spans="1:4" x14ac:dyDescent="0.2">
          <c r="A4" t="s">
            <v>8</v>
          </c>
          <c r="B4" t="s">
            <v>9</v>
          </c>
          <c r="C4" t="s">
            <v>10</v>
          </c>
          <c r="D4" t="s">
            <v>11</v>
          </c>
        </row>
        <row r="5" spans="1:4" x14ac:dyDescent="0.2">
          <c r="A5" s="1">
            <v>41014</v>
          </c>
          <c r="B5" t="str">
            <f>"txt"&amp;" fn"</f>
            <v>txt fn</v>
          </c>
          <c r="C5" t="str">
            <f>IF(1=2, "nope", "yep")</f>
            <v>yep</v>
          </c>
          <c r="D5">
            <v>84939202398</v>
          </c>
        </row>
      </sheetData>
      <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
    </worksheet>
    """
  end

  def shared_strings do
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="12" uniqueCount="12">
      <si>
        <t>A one</t>
      </si>
      <si>
        <t>B one</t>
      </si>
      <si>
        <t>C one</t>
      </si>
      <si>
        <t>D one</t>
      </si>
      <si>
        <t>A two</t>
      </si>
      <si>
        <t>B two</t>
      </si>
      <si>
        <t>C two</t>
      </si>
      <si>
        <t>D two</t>
      </si>
      <si>
        <t xml:space="preserve">The </t>
      </si>
      <si>
        <t>fox</t>
      </si>
      <si>
        <t>jumped over</t>
      </si>
      <si>
        <t>the fence</t>
      </si>
    </sst>
    """
  end

end