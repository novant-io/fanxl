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

  override Str toStr() { "${sheetId}:${name}" }
}