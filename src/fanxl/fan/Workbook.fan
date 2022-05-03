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

  ** The sheets for this workbook.
  Sheet[] sheets := [,]
}