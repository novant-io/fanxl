//
// Copyright (c) 2025, Novant LLC
// Licensed under the MIT License
//
// History:
//   27 Mar 2025  Andy Frank  Creation
//

*************************************************************************
** SheetTest
*************************************************************************

class SheetTest : AbstractWorkbookTest
{

//////////////////////////////////////////////////////////////////////////
// Access
//////////////////////////////////////////////////////////////////////////

  Void testCell()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet

    verifyEq(sh.lastRef, "F10")

    // sheet.cell
    verifyCell(sh.cell("A1"),  "0,0")
    verifyCell(sh.cell("D5"),  "3,4")
    verifyCell(sh.cell("Z5"),  null)
    verifyCell(sh.cell("D20"), null)

    // sheet.row.cell
    verifyCell(sh.row(0).cell(0),  "0,0")
    verifyCell(sh.row(4).cell(3),  "3,4")
    verifyCell(sh.row(4).cell(25), null)
    verifyEq(sh.row(19), null)
  }

//////////////////////////////////////////////////////////////////////////
// Update
//////////////////////////////////////////////////////////////////////////

  Void testUpdateCell()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet

    // sheet.update
    sh.updateCell("A1", "foo")
    verifyCell(sh.cell("A1"), "foo")
    verifyCell(sh.row(0).cell(0), "foo")

    // sheet.update
    sh.row(0).updateCell(0, "bar")
    verifyCell(sh.cell("A1"), "bar")
    verifyCell(sh.row(0).cell(0), "bar")
  }

  Void testUpdateCelsl()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet

    // sheet.updateCells
    sh.updateCells(0, ["foo","bar","zar"])
    verifyCell(sh.cell("A1"), "foo")
    verifyCell(sh.cell("B1"), "bar")
    verifyCell(sh.cell("C1"), "zar")
    verifyCell(sh.cell("D1"), "3,0")
    verifyCell(sh.cell("E1"), "4,0")

    // // with offset
    sh.updateCells(0, ["xxx","yyy","zzz"], 2)
    verifyCell(sh.cell("A1"), "foo")
    verifyCell(sh.cell("B1"), "bar")
    verifyCell(sh.cell("C1"), "xxx")
    verifyCell(sh.cell("D1"), "yyy")
    verifyCell(sh.cell("E1"), "zzz")
  }
}


