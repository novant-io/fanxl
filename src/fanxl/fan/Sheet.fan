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
@Js class Sheet
{

//////////////////////////////////////////////////////////////////////////
// Construction
//////////////////////////////////////////////////////////////////////////

  ** It-block ctor.
  internal new make(|This| f)
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

  // TODO: do we standardize on ref/alpha of index? or both with Obj?

  @NoDoc Int colWidth(Int col)
  {
    cwidths[col] ?: 10
  }

  ** Set the given column width by reference (ex: "A")
  @NoDoc Void setColWidth(Int col, Int width)
  {
    cwidths[col] = width
  }

  ** Get the given cell value (ex: "A5") or 'null' if not found.
  SheetCell? cell(Str ref)
  {
    cix := Util.cellRefToColIndex(ref)
    rix := Util.cellRefToRowIndex(ref)
    row := rows.getSafe(rix)
    if (row == null) return null
    return row.cell(cix)
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
    col := Util.colIndexToRef(numCols-1)
    row := this.numRows
    return "${col}${row}"
  }

//////////////////////////////////////////////////////////////////////////
// Rows
//////////////////////////////////////////////////////////////////////////

  ** Get the number of columns in this sheet.
  Int numCols()
  {
    // TODO FIXT
    ncols := 1
    rows.each |r|
    {
      ncols = ncols.max(r.size)
    }
    return ncols
  }

  ** Get the number of rows in this sheet.
  Int numRows() { rows.size }

  ** Get the row at the given index or 'null' if not found.
  SheetRow? row(Int index)
  {
    rows.getSafe(index)
  }

  ** Iterate the rows in this sheet.
  Void eachRow(|SheetRow,Int| f)
  {
    rows.each(f)
  }

  ** Iterate the rows in the given range for this sheet.
  Void eachRowRange(Range r, |SheetRow,Int| f)
  {
    rows.eachRange(r, f)
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

  ** Expand the sheet to the given row index
  private Void expandRows(Int index)
  {
    while (index >= rows.size) this.addRow
  }

  // TODO: not sure how this works yet
  internal Void _addRow(SheetRow row) { this.rows.add(row) }

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
      row := r.mapCells |c| { c.val }
      csv.writeRow(row)
    }
  }

  ** Trim trailing empty rows.
  internal Void trim()
  {
    while (rows.size > 0 && rows.last.isEmpty == true) rows.pop
  }

//////////////////////////////////////////////////////////////////////////
// Fields
//////////////////////////////////////////////////////////////////////////

  private Int:Int cwidths := Int:Int[:]    // column widths (or null for n/a)
  private SheetRow[] rows := SheetRow[,]   // rows for this sheet
}