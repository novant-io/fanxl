//
// Copyright (c) 2025, Novant LLC
// Licensed under the MIT License
//
// History:
//   4 Mar 2025  Andy Frank  Creation
//

*************************************************************************
** CsvWorkbookTest
*************************************************************************

class CsvWorkbookTest : Test
{

//////////////////////////////////////////////////////////////////////////
// test1
//////////////////////////////////////////////////////////////////////////

  Void test1()
  {
    wb := Fanxl.read(getTestFile("test_1.csv"))
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

  // -- CSV export round-trip
  Void test2()
  {
    wb := Fanxl.read(getTestFile("test_1.csv"))
    sh := wb.sheet

    // export to csv string
    csv := sh.toCsvStr

    // read back
    f := tempDir + `export_test.csv`
    f.out.print(csv).close
    wb2 := Fanxl.read(f)
    sh2 := wb2.sheet

    verifyEq(sh2.numRows, 10)
    sh2.eachRow |row,ri|
    {
      verifyEq(row.size, 6)
      row.eachCell |cell,ci|
      {
        verifyEq(cell.val, "${ci},${ri}")
      }
    }
  }

//////////////////////////////////////////////////////////////////////////
// test3
//////////////////////////////////////////////////////////////////////////

  // -- CSV round-trip should not produce trailing empty columns
  Void test3()
  {
    // build a sheet with jagged rows
    w := Workbook {}
    s := w.addSheet("Test")
    s.updateCell("A1", "Name")
    s.updateCell("B1", "Age")
    s.updateCell("C1", "City")
    s.updateCell("A2", "Alice")
    s.updateCell("B2", "30")
    // row 2 has no C column

    // write to CSV and read back
    f := tempDir + `csv_trailing_test.csv`
    out := f.out
    CsvWriter.write(s, out)
    out.close

    // read back and verify no empty column names
    wb := Fanxl.read(f)
    sh := wb.sheet
    verifyEq(sh.numRows, 2)
    sh.row(0).eachCell |c| { verify(!c.val.isEmpty) }
  }

//////////////////////////////////////////////////////////////////////////
// Support
//////////////////////////////////////////////////////////////////////////

  private File getTestFile(Str name)
  {
    file := Env.cur.workDir + `src/fanxl/test-csv/${name}`
    if (!file.exists) throw IOErr("File not found: ${file.osPath}")
    return file
  }
}


