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
    verifyEq(wb.sheets.size, 1)

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
// Support
//////////////////////////////////////////////////////////////////////////

  private File getTestFile(Str name)
  {
    file := Env.cur.workDir + `src/fanxl/test-csv/${name}`
    if (!file.exists) throw IOErr("File not found: ${file.osPath}")
    return file
  }
}


