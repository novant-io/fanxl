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

    sh := wb.sheets.first
    verifyEq(sh.rows.size, 10)
    sh.rows.each |row,ri|
    {
      verifyEq(row.cells.size, 6)
      row.cells.each |cell,ci|
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

    sh := wb.sheets.first
    verifyEq(sh.rows.size, 10)
    verifyEq(sh.rows[0].cells.join(";"), "0,0;1,0;2,0;3,0;4,0;5,0")
// TODO: do not add trailing empty cells?
    // verifyEq(sh.rows[1].cells.join(";"), "0,1;1,1;2,1;3,1;4,1;")
    verifyEq(sh.rows[1].cells.join(";"), "0,1;1,1;2,1;3,1;4,1")
    verifyEq(sh.rows[2].cells.join(";"), ";1,2;2,2;3,2;4,2;5,2")
    verifyEq(sh.rows[3].cells.join(";"), "0,3;;;3,3;4,3;5,3")
    verifyEq(sh.rows[4].cells.join(";"), "0,4;1,4;2,4;3,4;4,4;5,4")
// TODO: do not add trailing empty cells?
    //verifyEq(sh.rows[4].cells.join(";"), "0,5;1,5;;;;")
    verifyEq(sh.rows[5].cells.join(";"), "0,5;1,5")
    verifyEq(sh.rows[6].cells.join(";"), "0,6;1,6;2,6;3,6;4,6;5,6")
    verifyEq(sh.rows[7].cells.join(";"), "0,7;;2,7;;4,7;5,7")
    verifyEq(sh.rows[8].cells.join(";"), "0,8;1,8;2,8;3,8;4,8;5,8")
    verifyEq(sh.rows[9].cells.join(";"), ";1,9;;3,9;;5,9")
  }

//////////////////////////////////////////////////////////////////////////
// test3
//////////////////////////////////////////////////////////////////////////

  Void test3()
  {
    // x := File.os("/Users/andy/Desktop/test_x.xlsx")
    // XlsReader.read(x)

    f := tempDir + `write_test_1.xlsx`

    wb := Workbook {}
    wb.addSheet("Foo")
    wb.addSheet("Bar")
    wb.addSheet("Car")
    XlsWriter(wb).write(f.out)
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


