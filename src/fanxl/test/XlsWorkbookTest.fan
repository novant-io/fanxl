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
    verifyEq(s.row(0).joinCells(";"), "Alpha;Beta;Gamma")
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
}


