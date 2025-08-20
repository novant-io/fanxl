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
  new make(|This| f)
  {
    f(this)
    this.relId = "rId${id}"
  }

  ** Sheet id.
  const Int id

  ** Relationship id of this sheet.
  const Str relId

  ** Name of this sheet.
  Str name

  ** State of this sheet.
  internal Str state := "visible"

  override Str toStr() { "${id}:${name}" }

//////////////////////////////////////////////////////////////////////////
// Cells
//////////////////////////////////////////////////////////////////////////

  ** Get the given cell value (ex: "A5") or 'null' if not found.
  SheetCell? cell(Str ref)
  {
    cix := Util.cellRefToColIndex(ref)
    rix := Util.cellRefToRowIndex(ref)
    row := rows.getSafe(rix)
    if (row == null) return null
    return row.cells.getSafe(cix)
  }

  ** Update the cell reference.
  This updateCell(Str ref, Str val)
  {
    cix := Util.cellRefToColIndex(ref)
    rix := Util.cellRefToRowIndex(ref)

    // backfill missing rows
    expandRows(rix)

    // update
    row := rows[rix]
    row.updateCell(cix, val)
    return this
  }

  ** Update all the cells in the given row.
  This updateCells(Int index, Str[] cells, Int col := 0)
  {
    // backfill missing rows
    expandRows(index)
    rows[index].updateCells(cells, col)
    return this
  }

  ** Get the last cell reference (last row and column).
  internal Str lastRef()
  {
    // TODO FIXIT
    row  := this.numRows
    cols := 1
    rows.each |r|
    {
      cols = cols.max(r.cells.size)
    }
    return ('A'+cols-1).toChar + "${row}"
  }

//////////////////////////////////////////////////////////////////////////
// Rows
//////////////////////////////////////////////////////////////////////////

  ** Get the number of rows in this sheet.
  Int numRows() { rows.size }

  ** Get the row at the given index or 'null' if not found.
  SheetRow? row(Int index)
  {
    rows.getSafe(index)
  }

  ** Iterate the rows in this sheet.
  Void eachRow(|SheetRow| f)
  {
    rows.each(f)
  }

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

// TODO: goes away
  ** Rows for this sheet.
  @NoDoc SheetRow[] rows := SheetRow[,]

  ** Expand the sheet to the given row index
  private Void expandRows(Int index)
  {
    while (index >= rows.size) this.addRow
  }

//////////////////////////////////////////////////////////////////////////
// Export
//////////////////////////////////////////////////////////////////////////

  ** Export this sheet content to a CSV string.
  Str toCsvStr()
  {
    buf := StrBuf()
    toCsv(buf.out)
    return buf.toStr
  }

  ** Export this sheet content to CSV on the given output stream.
  Void toCsv(OutStream out)
  {
    csv := CsvOutStream(out)
    this.rows.each |r|
    {
      row := r.cells.map |c| { c.val }
      csv.writeRow(row)
    }
  }
}