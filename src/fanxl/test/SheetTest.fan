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

  Void testUpdateCells()
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

//////////////////////////////////////////////////////////////////////////
// NumCols / NumRows
//////////////////////////////////////////////////////////////////////////

  Void testNumColsRows()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet
    verifyEq(sh.numCols, 6)
    verifyEq(sh.numRows, 10)

    // update cell past current bounds
    sh.updateCell("H1", "wide")
    verifyEq(sh.numCols, 8)

    // add row expands numRows
    sh.addRow
    verifyEq(sh.numRows, 11)
  }

//////////////////////////////////////////////////////////////////////////
// AddRow
//////////////////////////////////////////////////////////////////////////

  Void testAddRow()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet
    verifyEq(sh.numRows, 10)

    row := sh.addRow
    verifyEq(sh.numRows, 11)
    verify(row.isEmpty)

    // write to new row and read back
    row.updateCell(0, "new")
    verifyEq(sh.row(10).cell(0).val, "new")
  }

//////////////////////////////////////////////////////////////////////////
// MapRows
//////////////////////////////////////////////////////////////////////////

  Void testMapRows()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet

    // mapRows skips header (row 0) and maps remaining rows
    vals := sh.mapRows |r,i| { r.cell(0).val }
    verifyEq(vals.size, 9)
    verifyEq(vals[0], "0,1")
    verifyEq(vals[8], "0,9")
  }

//////////////////////////////////////////////////////////////////////////
// EachRowRange
//////////////////////////////////////////////////////////////////////////

  Void testEachRowRange()
  {
    wb := Fanxl.read(getTestFile("test_1.xlsx"))
    sh := wb.sheet

    acc := Str[,]
    sh.eachRowRange(2..4) |r,i|
    {
      acc.add(r.cell(0).val)
    }
    verifyEq(acc.size, 3)
    verifyEq(acc[0], "0,2")
    verifyEq(acc[1], "0,3")
    verifyEq(acc[2], "0,4")
  }
}


