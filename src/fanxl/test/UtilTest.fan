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

  Void testRowIndex()
  {
    verifyEq(Util.cellRefToRowIndex("A1"),     0)
    verifyEq(Util.cellRefToRowIndex("B2"),     1)
    verifyEq(Util.cellRefToRowIndex("C35"),   34)
    verifyEq(Util.cellRefToRowIndex("Z200"), 199)
  }

  Void testColRef()
  {
    // single
    verifyEq(Util.colIndexToRef(0),  "A")
    verifyEq(Util.colIndexToRef(1),  "B")
    verifyEq(Util.colIndexToRef(2),  "C")
    verifyEq(Util.colIndexToRef(23), "X")
    verifyEq(Util.colIndexToRef(24), "Y")
    verifyEq(Util.colIndexToRef(25), "Z")

    // double
    verifyEq(Util.colIndexToRef(26), "AA")
    verifyEq(Util.colIndexToRef(27), "AB")
    verifyEq(Util.colIndexToRef(28), "AC")
    verifyEq(Util.colIndexToRef(49), "AX")
    verifyEq(Util.colIndexToRef(50), "AY")
    verifyEq(Util.colIndexToRef(51), "AZ")

    // big
    verifyEq(Util.colIndexToRef(135), "EF")
    verifyEq(Util.colIndexToRef(490), "RW")
    verifyEq(Util.colIndexToRef(701), "ZZ")
  }
}
