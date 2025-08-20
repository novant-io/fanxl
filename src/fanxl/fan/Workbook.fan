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
class Workbook
{
  ** It-block ctor.
  new make(|This| f) { f(this) }

  ** Get sheet with given 'name' or 'null' if none found.
  Sheet? sheet(Str name)
  {
    sheets.find |s| { s.name == name }
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

  internal Sheet[] sheets  := [,]   // sheets for this book
  private Int nextSheetId := 1     // next sheet_id to assign
}