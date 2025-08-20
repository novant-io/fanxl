//
// Copyright (c) 2025, Novant LLC
// Licensed under the MIT License
//
// History:
//   4 Mar 2025  Andy Frank  Creation
//

using util

**
** CsvReader reads and parses CSV files.
**
internal class CsvReader
{
  ** Read the given file and return a `Workbook` instance.
  static Workbook read(File file)
  {
    CsvInStream? in
    try
    {
      // init default sheet
      sheet := Sheet {
        // it.relId   = "csv"
        // it.sheetId = "csv"
        it.name    = "csv"
        it.id      = 1
      }

      // open file and interate rows
      in = CsvInStream(file.in)
      in.eachRow |vals|
      {
        // init row
        row := SheetRow {
          it.sheet = sheet
          it.index = 0
        }

        // add values
        vals.each |val| { row.cells.add(SheetCell { it.val=val }) }

        // add row
        sheet.rows.add(row)
      }

      // package in workbook and return
      book := Workbook {}
      book._addSheet(sheet)
      return book
    }
    catch (Err err)
    {
      throw ParseErr("Failed to parse spreadsheet", err)
    }
    finally { in?.close }
  }
}