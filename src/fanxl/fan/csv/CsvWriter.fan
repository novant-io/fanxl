//
// Copyright (c) 2025, Novant LLC
// Licensed under the MIT License
//
// History:
//   14 Mar 2025  Andy Frank  Creation
//

using util

**
** CsvWriter writes a Sheet to CSV format.
**
@Js internal class CsvWriter
{
  ** Write the given sheet to the provided output stream as CSV.
  static Void write(Sheet sheet, OutStream out)
  {
    csv := CsvOutStream(out)
    sheet.eachRow |row|
    {
      csv.writeRow(row.mapCells |c| { c.val })
    }
  }
}
