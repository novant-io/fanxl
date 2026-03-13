//
// Copyright (c) 2026, Novant LLC
// Licensed under the MIT License
//
// History:
//   13 Mar 2026  Andy Frank  Creation
//

*************************************************************************
** SheetRowTest
*************************************************************************

class SheetRowTest : AbstractWorkbookTest
{

//////////////////////////////////////////////////////////////////////////
// ToMap
//////////////////////////////////////////////////////////////////////////

  Void testToMap()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet

    // row(1) maps against row(0) headers
    map := sh.row(1).toMap
    verifyEq(map.size, 6)
    verifyEq(map["0,0"], "0,1")
    verifyEq(map["1,0"], "1,1")
    verifyEq(map["2,0"], "2,1")
    verifyEq(map["3,0"], "3,1")
    verifyEq(map["4,0"], "4,1")
    verifyEq(map["5,0"], "5,1")
  }

//////////////////////////////////////////////////////////////////////////
// IsEmpty
//////////////////////////////////////////////////////////////////////////

  Void testIsEmpty()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet

    // existing row with data is not empty
    verifyEq(sh.row(0).isEmpty, false)

    // new row with no cells is empty
    row := sh.addRow
    verifyEq(row.isEmpty, true)

    // adding empty-string cell is still empty
    row.updateCell(0, "")
    verifyEq(row.isEmpty, true)

    // adding non-empty cell makes it non-empty
    row.updateCell(0, "x")
    verifyEq(row.isEmpty, false)
  }
}
