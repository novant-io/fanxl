//
// Copyright (c) 2025, Novant LLC
// Licensed under the MIT License
//
// History:
//   27 Mar 2025  Andy Frank  Creation
//

*************************************************************************
** AbstractWorkbookTest
*************************************************************************

abstract class AbstractWorkbookTest : Test
{
  ** Load test file disk.
  protected File getTestFile(Str name)
  {
    file := Env.cur.workDir + `src/fanxl/test-xls/${name}`
    if (!file.exists) throw IOErr("File not found: ${file.osPath}")
    return file
  }

  ** Verify the given cell value.
  protected Void verifyCell(SheetCell? cell, Obj? expected)
  {
    if (expected == null) return verifyNull(cell)
    if (cell == null)     return fail("Cell is null")
    if (expected is Str)  return verifyEq(cell.val, expected)
    fail
  }
}


