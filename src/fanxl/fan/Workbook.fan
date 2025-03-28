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
    // TODO FIXIT relId/sheetId
    s := Sheet {
      it.name = "Sheet1"
      it.relId = "rId1"
      it.sheetId = "1"
    }
    sheets.add(s)
    return s
  }

// TODO GOES AWAY
  ** The sheets for this workbook.
  @NoDoc Sheet[] sheets := [,]
}