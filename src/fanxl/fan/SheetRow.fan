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
** SheetRow models a single row in a 'Sheet'.
**
class SheetRow
{
  ** It-block ctor.
  new make(|This| f) { f(this) }

  ** Index of this row in parent sheet.
  Int index

  ** Cells for this rows.
  SheetCell[] cells := [,]

  ** A row is empty if `cells` is empty for every value for `cells` is empty.
  Bool isEmpty()
  {
    if (cells.isEmpty) return true
    return cells.all |c| { c.val === "" }
  }

  override Str toStr() { cells.join(", ") }
}