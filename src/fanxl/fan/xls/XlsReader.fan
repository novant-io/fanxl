//
// Copyright (c) 2025, Novant LLC
// Licensed under the MIT License
//
// History:
//   3 Mar 2025  Andy Frank  Creation
//

using util
using xml

**
** XlsReader reads and parses XLSX files.
**
internal class XlsReader
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
    if (file == null) return
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
        // it.relId   = x.attr("id").val
        // it.sheetId = x.attr("sheetId").val
        it.name    = x.attr("name").val
        it.id      = x.attr("sheetId").val.toInt
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
// TODO FIXIT -> this will not work!  We need to load workbook.xml.refs to find rId:sheet<x>.xml
// <Relationship Id="rId3" Target="worksheets/sheet3.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
        sheetId := file.name[5..-5].toInt
        sheet   := book.sheets.find |s| { s.id == sheetId }
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
            val  := sst[sid] ?: "" // TODO
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
    while (sheet.rows.last?.isEmpty == true) sheet.rows.pop
  }
}