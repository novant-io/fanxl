//
// Copyright (c) 2022, Novant LLC
// Licensed under the MIT License
//
// History:
//   3 May 2022  Andy Frank  Creation
//

using util
using xml

**
** SheetCell models a single cell in a 'SheetRow'.
**
@Js class SheetCell
{
  ** It-block ctor.
  new make(|This| f) { f(this) }

  ** Cell value
  Str val := ""

  ** Get cell value as 'Float' or return 'null' if cell was empty.
  Float? float()
  {
    if (val.isEmpty) return null
    return val.toFloat
  }

  ** Get cell value as 'Time'
  Time? time()
  {
    if (val.isEmpty) return null
    fval  := val.toFloat
    frac  := fval - fval.floor
    fan   := DateTime.defVal.ticks + (24hr.ticks.toFloat * frac).toInt
    ticks := correctTicks(fan)
    return DateTime.makeTicks(ticks, TimeZone.utc).time
  }

  ** Get cell value as 'DateTime'
  DateTime? datetime(TimeZone tz := TimeZone.utc)
  {
    if (val.isEmpty) return null
    ser   := val.toFloat - 39448f
    orig  := Date(2008, Month.jan, 1).midnight(tz)
    days  := Duration("${ser}day")
    fan   := orig + days
    ticks := correctTicks(fan.ticks)
    return DateTime.makeTicks(ticks, fan.tz)
  }

  override Str toStr() { val }

  private Int correctTicks(Int orig)
  {
    // this convesion introduces some rounding errors
    // at the nanosecond level, so round to nearest
    // whole second
    floor := (orig / oneSecTicks) * oneSecTicks
    rem   := orig % oneSecTicks
    if (rem > 500ms.ticks) floor += 1sec.ticks
    return floor
  }

  private static const Int oneSecTicks := 1sec.ticks
}