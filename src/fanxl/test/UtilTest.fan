//
// Copyright (c) 2024, Novant LLC
// Licensed under the MIT License
//
// History:
//   6 Dec 2024  Andy Frank  Creation
//

*************************************************************************
** UtilTest
*************************************************************************

class UtilTest : Test
{
  Void testColIndex()
  {
    // single
    verifyEq(Util.cellRefToColIndex("A1"),    0)
    verifyEq(Util.cellRefToColIndex("B2"),    1)
    verifyEq(Util.cellRefToColIndex("C35"),   2)
    verifyEq(Util.cellRefToColIndex("Z200"), 25)
    // double
    verifyEq(Util.cellRefToColIndex("AA1"),    26)
    verifyEq(Util.cellRefToColIndex("AB1"),    27)
    verifyEq(Util.cellRefToColIndex("AC1"),    28)
    verifyEq(Util.cellRefToColIndex("BB1"),    53)
    verifyEq(Util.cellRefToColIndex("EF20"),  135)
    verifyEq(Util.cellRefToColIndex("RW35"),  490)
    verifyEq(Util.cellRefToColIndex("ZZ92"),  701)
  }
}
