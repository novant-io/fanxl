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
** SheetCell models a single cell in a 'SheetRow'.
**
class SheetCell
{
  ** It-block ctor.
  new make(|This| f) { f(this) }

  ** Cell value
  Str val

  override Str toStr() { val }
}