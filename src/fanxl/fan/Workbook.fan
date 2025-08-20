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
** Workbook is a document that models a list of spreadsheet 'Sheets'.
**
@Js class Workbook
{
  ** It-block ctor.
  new make(|This| f) { f(this) }

  ** Get the number of sheets in this workbook.
  Int numSheets() { sheets.size }

  ** Get sheet with given 'name' or 'null' if none found, or if
  ** no name is provied, return the first sheet in thsi workbook.
  Sheet? sheet(Str? name := null)
  {
    name == null
      ? sheets.first
      : sheets.find |s| { s.name == name }
  }

  // TODO: not sure how this works yet...
  @NoDoc Sheet? sheetAt(Int index)
  {
    sheets.getSafe(index)
  }

  ** Add a new sheet to this workbook.
  Sheet addSheet(Str name)
  {
    id := nextSheetId++
    s := Sheet {
      it.name = name
      it.id = id
      // it.relId = "rId${id}"
      // it.sheetId = "${id}"
    }
    // TODO: must have one row/cell
    s.updateCell("A1", "")
    sheets.add(s)
    return s
  }

  ** Iterate sheets in this workbook.
  Void eachSheet(|Sheet| f) { sheets.each(f) }

  // TODO: not sure how this works yet
  internal Void _addSheet(Sheet s) { sheets.add(s) }
  internal Sheet? _sheetById(Int id) { sheets.find |s| { s.id == id }}

  private Sheet[] sheets  := [,]   // sheets for this book
  private Int nextSheetId := 1     // next sheet_id to assign
}