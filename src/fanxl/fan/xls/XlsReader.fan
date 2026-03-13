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
      sheetFiles := readWorkbook(zip, book)
      readSheets(zip, book, sst, sheetFiles)

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
      first := k.elems.first
      if (first == null)
      {
        // empty <si/> entry
        sst.add(i, "")
      }
      else if (first.name == "t")
      {
        // simple string: <si><t>text</t></si>
        s := first.children.first as XText
        sst.add(i, s?.val ?: "")
      }
      else
      {
        // rich text: <si><r><rPr>...</rPr><t>text</t></r>...</si>
        buf := StrBuf()
        k.elems.each |r|
        {
          t := r.elems.find |e| { e.name == "t" }
          s := t?.children?.first as XText
          if (s != null) buf.add(s.val)
        }
        sst.add(i, buf.toStr)
      }
    }
  }

  ** Parse 'workbook.xml' and return map of sheetId to sheet file path.
  private static Int:Str readWorkbook(Zip zip, Workbook book)
  {
    // read workbook rels to map rId -> target path
    rels := Str:Str[:]
    relsFile := zip.contents[`/xl/_rels/workbook.xml.rels`]
    if (relsFile != null)
    {
      rdoc := XParser(relsFile.in).parseDoc
      rdoc.root.elems.each |r|
      {
        id     := r.attr("Id", false)?.val
        target := r.attr("Target", false)?.val
        if (id != null && target != null) rels[id] = target
      }
    }

    // read workbook
    file := zip.contents[`/xl/workbook.xml`]
    doc  := XParser(file.in).parseDoc
    root := doc.root

    sheetFiles := Int:Str[:]
    sheets := root.elems.find |k| { k.name == "sheets" }
    sheets.elems.each |x|
    {
      sheetId := x.attr("sheetId").val.toInt

      book._addSheet(Sheet {
        it.name    = x.attr("name").val
        it.id      = sheetId
        it.state   = x.attr("state", false)?.val ?: "visible"
      })

      // resolve sheet file path via r:id -> rels
      rId := x.attr("id", false)?.val
      if (rId != null && rels.containsKey(rId))
        sheetFiles[sheetId] = "/xl/" + rels[rId]
    }

    return sheetFiles
  }

  ** Read sheet data using resolved file mappings.
  private static Void readSheets(Zip zip, Workbook book, Int:Str sst, Int:Str sheetFiles)
  {
    sheetFiles.each |path, sheetId|
    {
      sheet := book._sheetById(sheetId)
      if (sheet == null) return
      file := zip.contents[path.toUri]
      if (file == null) return
      readSheet(file, sheet, sst)
    }
  }

  ** Parse 'sheet<x>.xml`
  private static Void readSheet(File file, Sheet sheet, Int:Str sst)
  {
    doc  := XParser(file.in).parseDoc
    root := doc.root

    data := root.elems.find |e| { e.name == "sheetData" }
    if (data == null) return
    lastRowNum := 0
    data.elems.each |xr,i|
    {
      rowNum := xr.attr("r", false)?.val?.toInt ?: (lastRowNum + 1)

      // backfill sparse/missing rows
      while (rowNum > lastRowNum + 1)
      {
        lastRowNum++
        sheet._addRow(SheetRow {
          it.sheet = sheet
          it.index = lastRowNum
        })
      }
      lastRowNum = rowNum

      row := SheetRow {
        it.sheet = sheet
        it.index = rowNum
      }
      // track last cix to backfill sparse cells
      lastcix := -1
      xr.elems.each |xc|
      {
        ref  := xc.attr("r", false)?.val
        type := xc.attr("t", false)?.val

        // backfill empty/sparse columns if needed
        cix  := ref != null ? Util.cellRefToColIndex(ref) : lastcix + 1
        miss := cix - lastcix
        while (miss-- > 1) row._addCell(SheetCell { it.val="" })
        lastcix = cix

        // read value based on type
        switch (type)
        {
          case "s":
            // shared string
            sid  := xc.elems.first.text.val.toInt
            val  := sst[sid] ?: "" // TODO
            cell := SheetCell { it.val=val }
            row._addCell(cell)

          case "b":
            // boolean: <v>0</v> or <v>1</v>
            bval := xc.elems.first?.text?.val
            cell := SheetCell { it.val = bval == "1" ? "true" : "false" }
            row._addCell(cell)

          case "inlineStr":
            // inline string: <is><t>value</t></is>
            val  := xc.elems.first?.elems?.first?.text?.val ?: ""
            cell := SheetCell { it.val=val }
            row._addCell(cell)

          default:
            val  := xc.elems.first?.text?.val ?: ""
            cell := SheetCell { it.val=val }
            row._addCell(cell)
        }
      }

      sheet._addRow(row)
    }

    // trim trailing empty rows
    sheet.trim
  }
}