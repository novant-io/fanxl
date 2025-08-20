//
// Copyright (c) 2024, Novant LLC
// Licensed under the MIT License
//
// History:
//   6 Dec 2024  Andy Frank  Creation
//

@Js internal const class Util
{
  ** Parse a Str reference such as 'C12' to the zero-based column index '2'.
  static Int cellRefToColIndex(Str ref)
  {
    cix := 0
    pos := 0
    while (ref[pos].isAlpha)
    {
      // convert to 1-based index
      v := ref[pos] - 'A' + 1
      cix = cix * 26 + v
      pos++
    }
    // convert back to 0-based
    return cix-1
  }

  ** Parse a Str reference such as 'C12' to the zero-based row index '11'.
  static Int cellRefToRowIndex(Str ref)
  {
    pos := 0
    while (ref[pos].isAlpha) pos++
    return ref[pos..-1].toInt - 1
  }
}