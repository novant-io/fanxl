//
// Copyright (c) 2026, Novant LLC
// Licensed under the MIT License
//
// History:
//   14 Mar 2026  Andy Frank  Creation
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
    // determine number of columns from header row
    header := sheet.row(0)
    numCols := 0
    if (header != null)
    {
      header.eachCell |c,i| { if (!c.val.isEmpty) numCols = i + 1 }
    }

    csv := CsvOutStream(out)
    sheet.eachRow |row|
    {
      vals := Str[,]
      numCols.times |i| { vals.add(row.cell(i)?.val ?: "") }
      csv.writeRow(vals)
    }
  }
}