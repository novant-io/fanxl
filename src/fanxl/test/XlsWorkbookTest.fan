//
// Copyright (c) 2024, Novant LLC
// Licensed under the MIT License
//
// History:
//   6 Nov 2024  Andy Frank  Creation
//

*************************************************************************
** XlsWorkbookTest
*************************************************************************

class XlsWorkbookTest : Test
{

//////////////////////////////////////////////////////////////////////////
// test1
//////////////////////////////////////////////////////////////////////////

  Void test1()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    verifyEq(wb.numSheets, 1)

    sh := wb.sheet
    verifyEq(sh.numRows, 10)
    sh.eachRow |row,ri|
    {
      verifyEq(row.size, 6)
      row.eachCell |cell,ci|
      {
        verifyEq(cell.val, "${ci},${ri}")
      }
    }
  }

//////////////////////////////////////////////////////////////////////////
// test2
//////////////////////////////////////////////////////////////////////////

  // -- Test sparse columns
  //
  //  0,0 | 1,0 | 2,0 | 3,0 |4,0 | 5,0
  //  0,1 | 1,1 | 2,1 | 3,1 |4,1 |
  //      | 1,2 | 2,2 | 3,2 |4,2 | 5,2
  //  0,3 |     |     | 3,3 |4,3 | 5,3
  //  0,4 | 1,4 | 2,4 | 3,4 |4,4 | 5,4
  //  0,5 | 1,5 |     |     |    |
  //  0,6 | 1,6 | 2,6 | 3,6 |4,6 | 5,6
  //  0,7 |     | 2,7 |     |4,7 | 5,7
  //  0,8 | 1,8 | 2,8 | 3,8 |4,8 | 5,8
  //      | 1,9 |     | 3,9 |    | 5,9

  Void test2()
  {
    wb := Fanxl.read(getTestFile("test_2.xlsx"))
    verifyEq(wb.numSheets, 1)

    sh := wb.sheet
    verifyEq(sh.numRows, 10)
    verifyEq(sh.row(0).joinCells(";"), "0,0;1,0;2,0;3,0;4,0;5,0")
// TODO: do not add trailing empty cells?
    // verifyEq(sh.rows[1].joinCells(";"), "0,1;1,1;2,1;3,1;4,1;")
    verifyEq(sh.row(1).joinCells(";"), "0,1;1,1;2,1;3,1;4,1")
    verifyEq(sh.row(2).joinCells(";"), ";1,2;2,2;3,2;4,2;5,2")
    verifyEq(sh.row(3).joinCells(";"), "0,3;;;3,3;4,3;5,3")
    verifyEq(sh.row(4).joinCells(";"), "0,4;1,4;2,4;3,4;4,4;5,4")
// TODO: do not add trailing empty cells?
    //verifyEq(sh.rows[4].joinCells(";"), "0,5;1,5;;;;")
    verifyEq(sh.row(5).joinCells(";"), "0,5;1,5")
    verifyEq(sh.row(6).joinCells(";"), "0,6;1,6;2,6;3,6;4,6;5,6")
    verifyEq(sh.row(7).joinCells(";"), "0,7;;2,7;;4,7;5,7")
    verifyEq(sh.row(8).joinCells(";"), "0,8;1,8;2,8;3,8;4,8;5,8")
    verifyEq(sh.row(9).joinCells(";"), ";1,9;;3,9;;5,9")
  }

//////////////////////////////////////////////////////////////////////////
// test3
//////////////////////////////////////////////////////////////////////////

  Void test3()
  {
    // temp file
    f := tempDir + `write_test_1.xlsx`

    // write file
    w := Workbook {}
    // foo
    s := w.addSheet("Foo")
    s.updateCell("A1", "Alpha")
    s.updateCell("B1", "Beta")
    s.updateCell("C1", "Gamma")
    // bar
    w.addSheet("Bar")
    // zar
    w.addSheet("Zar")
    XlsWriter(w).write(f.out)

    // read back file
    r := XlsReader.read(f)
    verifyEq(r.numSheets, 3)
    s1 := r.sheet("Foo")
    verifyEq(s1.row(0).joinCells(";"), "Alpha;Beta;Gamma")
  }

//////////////////////////////////////////////////////////////////////////
// test4
//////////////////////////////////////////////////////////////////////////

  // -- Multi-sheet round-trip with data on each sheet
  Void test4()
  {
    f := tempDir + `write_test_4.xlsx`

    // write
    w := Workbook {}
    s1 := w.addSheet("Sheet1")
    s1.updateCell("A1", "Name")
    s1.updateCell("B1", "Age")
    s1.updateCell("A2", "Alice")
    s1.updateCell("B2", "30")
    s1.updateCell("A3", "Bob")
    s1.updateCell("B3", "25")

    s2 := w.addSheet("Sheet2")
    s2.updateCell("A1", "X")
    s2.updateCell("B1", "Y")
    s2.updateCell("C1", "Z")

    s3 := w.addSheet("Sheet3")
    s3.updateCell("A1", "Solo")

    XlsWriter(w).write(f.out)

    // read back
    r := XlsReader.read(f)
    verifyEq(r.numSheets, 3)

    // verify sheet1
    rs1 := r.sheet("Sheet1")
    verifyEq(rs1.numRows, 3)
    verifyEq(rs1.row(0).joinCells(";"), "Name;Age")
    verifyEq(rs1.row(1).joinCells(";"), "Alice;30")
    verifyEq(rs1.row(2).joinCells(";"), "Bob;25")

    // verify sheet2
    rs2 := r.sheet("Sheet2")
    verifyEq(rs2.numRows, 1)
    verifyEq(rs2.row(0).joinCells(";"), "X;Y;Z")

    // verify sheet3
    rs3 := r.sheet("Sheet3")
    verifyEq(rs3.numRows, 1)
    verifyEq(rs3.row(0).joinCells(";"), "Solo")
  }

//////////////////////////////////////////////////////////////////////////
// test5
//////////////////////////////////////////////////////////////////////////

  // -- XML special characters in sheet names and cell values
  Void test5()
  {
    f := tempDir + `write_test_5.xlsx`

    // write
    w := Workbook {}
    s := w.addSheet("Sales & Returns")
    s.updateCell("A1", "Tom & Jerry")
    s.updateCell("B1", "x < y")
    s.updateCell("C1", "a > b")
    s.updateCell("A2", "she said \"hi\"")
    s.updateCell("B2", "it's fine")
    XlsWriter(w).write(f.out)

    // read back
    r := XlsReader.read(f)
    rs := r.sheet("Sales & Returns")
    verifyNotNull(rs)
    verifyEq(rs.row(0).joinCells(";"), "Tom & Jerry;x < y;a > b")
    verifyEq(rs.row(1).cell(0).val, "she said \"hi\"")
    verifyEq(rs.row(1).cell(1).val, "it's fine")
  }

//////////////////////////////////////////////////////////////////////////
// test6
//////////////////////////////////////////////////////////////////////////

  // -- Many columns past Z (AA, AB, etc.)
  Void test6()
  {
    f := tempDir + `write_test_6.xlsx`

    // write 30 columns
    w := Workbook {}
    s := w.addSheet("Wide")
    30.times |i|
    {
      col := Util.colIndexToRef(i)
      s.updateCell("${col}1", "c${i}")
    }
    XlsWriter(w).write(f.out)

    // read back
    r := XlsReader.read(f)
    rs := r.sheet("Wide")
    verifyEq(rs.numCols, 30)
    verifyEq(rs.row(0).cell(0).val,  "c0")   // A
    verifyEq(rs.row(0).cell(25).val, "c25")  // Z
    verifyEq(rs.row(0).cell(26).val, "c26")  // AA
    verifyEq(rs.row(0).cell(27).val, "c27")  // AB
    verifyEq(rs.row(0).cell(29).val, "c29")  // AD
  }

//////////////////////////////////////////////////////////////////////////
// test7
//////////////////////////////////////////////////////////////////////////

  // -- Sparse data round-trip
  Void test7()
  {
    f := tempDir + `write_test_7.xlsx`

    // write with gaps: A1, C1, F1 (skip B, D, E)
    w := Workbook {}
    s := w.addSheet("Sparse")
    s.updateCell("A1", "one")
    s.updateCell("C1", "three")
    s.updateCell("F1", "six")
    // row 3 with no row 2 data
    s.updateCell("B3", "mid")
    XlsWriter(w).write(f.out)

    // read back
    r := XlsReader.read(f)
    rs := r.sheet("Sparse")
    verifyEq(rs.row(0).cell(0).val, "one")
    verifyEq(rs.row(0).cell(1).val, "")
    verifyEq(rs.row(0).cell(2).val, "three")
    verifyEq(rs.row(0).cell(3).val, "")
    verifyEq(rs.row(0).cell(4).val, "")
    verifyEq(rs.row(0).cell(5).val, "six")
  }

//////////////////////////////////////////////////////////////////////////
// test8
//////////////////////////////////////////////////////////////////////////

  // -- Reader handles rows and cells with missing r attributes
  Void test8()
  {
    f := tempDir + `reader_test_8.xlsx`
    sheetXml :=
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
        <sheetData>
          <row>
            <c t=\"inlineStr\"><is><t>A</t></is></c>
            <c t=\"inlineStr\"><is><t>B</t></is></c>
            <c t=\"inlineStr\"><is><t>C</t></is></c>
          </row>
          <row>
            <c t=\"inlineStr\"><is><t>D</t></is></c>
            <c t=\"inlineStr\"><is><t>E</t></is></c>
          </row>
        </sheetData>
      </worksheet>"
    writeTestXlsx(f, sheetXml)

    r := XlsReader.read(f)
    sh := r.sheet
    verifyEq(sh.numRows, 2)
    verifyEq(sh.row(0).joinCells(";"), "A;B;C")
    verifyEq(sh.row(1).joinCells(";"), "D;E")
  }

//////////////////////////////////////////////////////////////////////////
// test9
//////////////////////////////////////////////////////////////////////////

  // -- Reader handles boolean cells
  Void test9()
  {
    f := tempDir + `reader_test_9.xlsx`
    sheetXml :=
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
        <sheetData>
          <row r=\"1\">
            <c r=\"A1\" t=\"b\"><v>1</v></c>
            <c r=\"B1\" t=\"b\"><v>0</v></c>
            <c r=\"C1\" t=\"inlineStr\"><is><t>text</t></is></c>
          </row>
        </sheetData>
      </worksheet>"
    writeTestXlsx(f, sheetXml)

    r := XlsReader.read(f)
    sh := r.sheet
    verifyEq(sh.numRows, 1)
    verifyEq(sh.row(0).cell(0).val, "true")
    verifyEq(sh.row(0).cell(1).val, "false")
    verifyEq(sh.row(0).cell(2).val, "text")
  }

//////////////////////////////////////////////////////////////////////////
// test10
//////////////////////////////////////////////////////////////////////////

  // -- Reader handles sheet with no sheetData element
  Void test10()
  {
    f := tempDir + `reader_test_10.xlsx`
    sheetXml :=
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
      </worksheet>"
    writeTestXlsx(f, sheetXml)

    r := XlsReader.read(f)
    sh := r.sheet
    verifyEq(sh.numRows, 0)
  }

//////////////////////////////////////////////////////////////////////////
// test11
//////////////////////////////////////////////////////////////////////////

  // -- Reader handles empty shared string entries
  Void test11()
  {
    f := tempDir + `reader_test_11.xlsx`
    sheetXml :=
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
        <sheetData>
          <row r=\"1\">
            <c r=\"A1\" t=\"s\"><v>0</v></c>
            <c r=\"B1\" t=\"s\"><v>1</v></c>
            <c r=\"C1\" t=\"s\"><v>2</v></c>
          </row>
        </sheetData>
      </worksheet>"
    sstXml :=
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
        <si><t>hello</t></si>
        <si/>
        <si><t>world</t></si>
      </sst>"
    writeTestXlsx(f, sheetXml, sstXml)

    r := XlsReader.read(f)
    sh := r.sheet
    verifyEq(sh.row(0).cell(0).val, "hello")
    verifyEq(sh.row(0).cell(1).val, "")
    verifyEq(sh.row(0).cell(2).val, "world")
  }

//////////////////////////////////////////////////////////////////////////
// Support
//////////////////////////////////////////////////////////////////////////

  private File getTestFile(Str name)
  {
    file := Env.cur.workDir + `src/fanxl/test-xls/${name}`
    if (!file.exists) throw IOErr("File not found: ${file.osPath}")
    return file
  }

  ** Write a minimal XLSX file with custom sheet and optional shared strings XML.
  private Void writeTestXlsx(File f, Str sheetXml, Str? sstXml := null)
  {
    zip := Zip.write(f.out)

    // workbook.xml
    out := zip.writeNext(`/xl/workbook.xml`)
    out.printLine(
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"
                xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
        <sheets>
          <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>
        </sheets>
      </workbook>")
    out.close

    // workbook.xml.rels
    out = zip.writeNext(`/xl/_rels/workbook.xml.rels`)
    out.printLine(
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
        <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
      </Relationships>")
    out.close

    // sheet1.xml
    out = zip.writeNext(`/xl/worksheets/sheet1.xml`)
    out.printLine(sheetXml)
    out.close

    // optional shared strings
    if (sstXml != null)
    {
      out = zip.writeNext(`/xl/sharedStrings.xml`)
      out.printLine(sstXml)
      out.close
    }

    zip.close
  }
}


