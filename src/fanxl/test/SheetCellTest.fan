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


