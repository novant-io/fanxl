//
// Copyright (c) 2024, Novant LLC
// Licensed under the MIT License
//
// History:
//   12 Nov 2024  Andy Frank  Creation
//

*************************************************************************
** SheetCellTest
*************************************************************************

class SheetCellTest : Test
{

//////////////////////////////////////////////////////////////////////////
// Time
//////////////////////////////////////////////////////////////////////////

  Void testTime()
  {
    // simple even tests
    verifyTime("1.00", Time(0,  0, 0))
    verifyTime("1.25", Time(6,  0, 0))
    verifyTime("1.50", Time(12, 0, 0))
    verifyTime("1.75", Time(18, 0, 0))

    // more tests
    verifyTime("1.020833333333333", Time( 0, 30, 0))
    verifyTime("1.0625",            Time( 1, 30, 0))
    verifyTime("1.104166666666667", Time( 2, 30, 0))
    verifyTime("1.145833333333333", Time( 3, 30, 0))
    verifyTime("1.1875",            Time( 4, 30, 0))
    verifyTime("1.229166666666667", Time( 5, 30, 0))
    verifyTime("1.854166666666667", Time(20, 30, 0))
    verifyTime("1.895833333333333", Time(21, 30, 0))
    verifyTime("1.9375",            Time(22, 30, 0))
    verifyTime("1.979166666666667", Time(23, 30, 0))
  }

  private Void verifyTime(Str val, Time t)
  {
    c := SheetCell { it.val=val }
    verifyEq(c.time, t)
  }

//////////////////////////////////////////////////////////////////////////
// DateTime
//////////////////////////////////////////////////////////////////////////

  Void testDateTime()
  {
    verifyDateTime("45607.666666666664", DateTime("2024-11-11T16:00:00Z UTC"))
    verifyDateTime("45607.708333333336", DateTime("2024-11-11T17:00:00Z UTC"))
    verifyDateTime("45607.75",           DateTime("2024-11-11T18:00:00Z UTC"))
    verifyDateTime("45608.125",          DateTime("2024-11-12T03:00:00Z UTC"))
    verifyDateTime("45608.166666666664", DateTime("2024-11-12T04:00:00Z UTC"))
    verifyDateTime("45608.208333333336", DateTime("2024-11-12T05:00:00Z UTC"))
  }

  private Void verifyDateTime(Str val, DateTime dt)
  {
    c := SheetCell { it.val=val }
    verifyEq(c.datetime, dt)
  }
}


