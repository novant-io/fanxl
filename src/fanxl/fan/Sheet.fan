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
** Sheet models a single worksheet in a 'Workbook'.
**
class Sheet
{

//////////////////////////////////////////////////////////////////////////
// Construction
//////////////////////////////////////////////////////////////////////////

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

  override Str toStr() { "${sheetId}:${name}" }

//////////////////////////////////////////////////////////////////////////
// Cells
//////////////////////////////////////////////////////////////////////////

  ** Get the given cell value (ex: "A5") or 'null' if not found.
  Str? cell(Str ref)
  {
    cix := Util.cellRefToColIndex(ref)
    rix := Util.cellRefToRowIndex(ref)
    row := rows.getSafe(rix)
    if (row == null) return null
    return row.cells.getSafe(cix)?.val
  }

  ** Update the cell reference.
  Void updateCell(Str ref, Str val)
  {
    cix := Util.cellRefToColIndex(ref)
    rix := Util.cellRefToRowIndex(ref)

    // backfill missing rows
    while (rix >= rows.size) this.addRow

    // update
    row := rows[rix]
    row.update(cix, val)
  }

//////////////////////////////////////////////////////////////////////////
// Rows
//////////////////////////////////////////////////////////////////////////

  ** Rows for this sheet.
  SheetRow[] rows := SheetRow[,]

  ** Add a new row to this sheet.
  SheetRow addRow()
  {
    row := SheetRow {
      it.sheet = this
      it.index = rows.size + 1
    }
    rows.add(row)
    return row
  }

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
}