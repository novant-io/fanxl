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
** Fanxl reads and parses XLSX files.
**
const class Fanxl
{
  ** Read the given file and return a `Workbook` instance.
  static Workbook read(File file)
  {
    Zip? zip
    try
    {
      book := Workbook {}
      sst  := Int:Str[:]

      zip = Zip.open(file)
      readSharedStrings(zip, sst)
      readWorkbook(zip, book)
      readSheets(zip, book, sst)

      return book
    }
    catch (Err err)
    {
      throw ParseErr("Failed to parse spreadsheet", err)
    }
    finally { zip?.close }
  }

  ** Parse 'sharedStrings.xml`
  private static Void readSharedStrings(Zip zip, Int:Str sst)
  {
    file := zip.contents[`/xl/sharedStrings.xml`]
    doc  := XParser(file.in).parseDoc
    root := doc.root
    root.elems.each |k,i|
    {
      t := k.elems.first
      s := t.children.first as XText
// TODO <si><r>...
if (s == null) return
      sst.add(i, s.val)
    }
  }

  ** Parse 'workbook.xml`
  private static Void readWorkbook(Zip zip, Workbook book)
  {
    file := zip.contents[`/xl/workbook.xml`]
    doc  := XParser(file.in).parseDoc
    root := doc.root
    // dump(root)

    sheets := root.elems.find |k| { k.name == "sheets" }
    sheets.elems.each |x|
    {
      book.sheets.add(Sheet {
        it.relId   = x.attr("id").val
        it.sheetId = x.attr("sheetId").val
        it.name    = x.attr("name").val
        it.state   = x.attr("state", false)?.val ?: "visible"
      })
    }
  }

  ** Read all `sheet<x>.xml` files.
  private static Void readSheets(Zip zip, Workbook book, Int:Str sst)
  {
    zip.contents.each |file|
    {
      if (file.name.startsWith("sheet") && file.ext == "xml")
      {
        sheetId := file.name[5..-5]
        sheet   := book.sheets.find |s| { s.sheetId == sheetId }
        //if (sheet == null) throw ArgErr("Sheet not found '${sheetId}'")
// TODO
if (sheet == null) return
        readSheet(file, sheet, sst)
      }
    }
  }

  ** Parse 'sheet<x>.xml`
  private static Void readSheet(File file, Sheet sheet, Int:Str sst)
  {
    doc  := XParser(file.in).parseDoc
    root := doc.root

    data := root.elems.find |e| { e.name == "sheetData" }
    data.elems.each |xr,i|
    {
      row := SheetRow {
        it.sheet = sheet
        it.index = xr.attr("r").val.toInt
      }
      // track last cix to backfill sparse cells
      lastcix := -1
      xr.elems.each |xc|
      {
        ref  := xc.attr("r").val
        type := xc.attr("t", false)?.val

        // backfill empty/sparse columns if needed
        cix  := Util.cellRefToColIndex(ref)
        miss := cix - lastcix
        while (miss-- > 1) row.cells.add(SheetCell { it.val="" })
        lastcix = cix

        // read value based on type
        switch (type)
        {
          case "s":
            // shared string
            sid  := xc.elems.first.text.val.toInt
            val  := sst[sid] ?: "X" // TODO
            cell := SheetCell { it.val=val }
            row.cells.add(cell)

          default:
            val  := xc.elems.first?.text?.val ?: ""
            cell := SheetCell { it.val=val }
            row.cells.add(cell)
        }
      }

      sheet.rows.add(row)
    }

    // trim trailing empty rows
    while (sheet.rows.last.isEmpty) sheet.rows.pop
  }
}