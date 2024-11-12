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
class SheetCell
{
  ** It-block ctor.
  new make(|This| f) { f(this) }

  ** Cell value
  Str val

  ** Get cell value as 'Float' or return 'null' if cell was empty.
  Float? float()
  {
    if (val.isEmpty) return null
    return val.toFloat
  }

  ** Get cell value as 'DateTime'
  DateTime? datetime(TimeZone tz := TimeZone.utc)
  {
    if (val.isEmpty) return null
    ser  := val.toFloat - 39448f
    orig := Date(2008, Month.jan, 1).midnight(tz)
    days := Duration("${ser}day")
    fan  := orig + days
    // TODO: round/clamp to nearest min
    return fan
  }

  override Str toStr() { val }
}