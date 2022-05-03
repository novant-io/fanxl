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
** Sheet models a single spreadsheet in a 'Workbook'.
**
class Sheet
{
  ** It-block ctor.
  new make(|This| f) { f(this) }

  ** Relationship id of this sheet.
  Str relId

  ** Internal id of this sheet.
  Str sheetId

  ** Name of this sheet.
  Str name

  ** State of this sheet.
  Str state := "visible"

  ** Rows for this sheet.
  SheetRow[] rows := SheetRow[,]

  ** Create a new list which is the result of calling 'f' for
  ** every row (excluding header) in this sheet.
  Obj[] mapRows(|SheetRow row, Int index->Obj| f)
  {
    acc := Obj[,]
    if (rows.size > 1)
    {
      rows.eachRange(1..-1) |r,i|
      {
        acc.add(f(r,i))
      }
    }
    return acc
  }

  ** Map the given row index to a 'Str:Str?' map where the map
  ** keys are the column values for `rows[0]`.
  Str:Str? rowToMap(Int row)
  {
    if (row == 0) throw ArgErr("Cannot map row 0")
    if (row >= rows.size) throw ArgErr("Invalid row: ${row}")

    cols := rows[0].cells
    vals := rows[row].cells

    map := Str:Str[:] { it.ordered=true }
    cols.each |c,i|
    {
      k := c.val
      v := vals.getSafe(i)?.val
      map[k] = v
    }

    return map
  }

  override Str toStr() { "${sheetId}:${name}" }
}