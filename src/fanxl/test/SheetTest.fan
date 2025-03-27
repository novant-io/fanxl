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
    sh := wb.sheets.first

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
  }
}


