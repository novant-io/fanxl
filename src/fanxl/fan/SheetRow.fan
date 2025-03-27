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

  ** Parent 'Sheet' instance.
  internal Sheet? sheet

  ** Index of this row in parent sheet.
  Int index

// TODO GOES AWAY

  ** Cells for this rows.
  SheetCell[] cells := [,]

  ** Get the given cell value or 'null' if not found.
  SheetCell? cell(Int col)
  {
    cells.getSafe(col)
  }

  ** Update the cell reference.
  Void updateCell(Int col, Str val)
  {
    // backfill missing cells
    while (col >= cells.size)
    {
      ix := cells.size
      cells.add(SheetCell {})
    }

    // update
    cells[col].val = val
  }

  ** Update a range of cells starting at given column.
  Void updateCells(Str[] vals, Int col := 0)
  {
    // expand row if needed
    maxc := col + vals.size
    while (maxc >= cells.size) cells.add(SheetCell {})

    // update
    vals.each |v,i|
    {
      cells[col+i].val = v
    }
  }

  ** A row is empty if `cells` is empty for every value
  ** for `cells` is empty.
  Bool isEmpty()
  {
    if (cells.isEmpty) return true
    return cells.all |c| { c.val === "" }
  }

  ** Map this row to a 'Str:Str?' map where the map keys are
  ** the column values for `rows[0]` from parent sheet.
  Str:Str? toMap()
  {
    cols := sheet.rows[0].cells
    vals := sheet.rows[index-1].cells

    map := Str:Str?[:] { it.ordered=true }
    cols.each |c,i|
    {
      k := c.val
      v := vals.getSafe(i)?.val
      map[k] = v
    }

    return map
  }

  override Str toStr() { cells.join(", ") }
}