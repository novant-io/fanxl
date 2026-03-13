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
  // Excel maximum limits for bounds checking
  private static const Int maxRows := 1_048_576
  private static const Int maxCols := 16_384

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
    xml := file.readAllStr

    // extract <worksheet ...> opening tag to preserve namespace declarations
    wsStart := xml.index("<worksheet")
    if (wsStart == null) return
    wsGt := xml.index(">", wsStart)
    wsTag := xml[wsStart..wsGt]
    if (wsTag.endsWith("/")) return  // degenerate self-closing worksheet
    wsSuffix := "</worksheet>"

    // find <sheetData> section
    sdStart := xml.index("<sheetData")
    if (sdStart == null) return
    sdGt := xml.index(">", sdStart)
    if (xml[sdGt-1] == '/') return  // self-closing <sheetData/>
    sdEnd := xml.index("</sheetData>")
    if (sdEnd == null) return

    pos := sdGt + 1
    lastRowNum := 0

    while (pos < sdEnd)
    {
      // find next <row element
      rs := xml.index("<row", pos)
      if (rs == null || rs >= sdEnd) break

      // verify tag name boundary (not <rowBreaks etc.)
      nc := xml[rs + 4]
      if (nc != ' ' && nc != '>' && nc != '/')
      {
        pos = rs + 4
        continue
      }

      // find close of this element
      gt := xml.index(">", rs)
      if (xml[gt-1] == '/')
      {
        // self-closing <row ... /> (empty row)
        rowXml := wsTag + xml[rs..gt] + wsSuffix
        pos = gt + 1
        xr := XParser(rowXml.in).parseDoc.root.elems.first
        rowNum := xr.attr("r", false)?.val?.toInt ?: (lastRowNum + 1)
        if (rowNum > maxRows) throw ParseErr("Row $rowNum exceeds max ($maxRows)")
        while (rowNum > lastRowNum + 1)
        {
          lastRowNum++
          sheet._addRow(SheetRow { it.sheet = sheet; it.index = lastRowNum })
        }
        lastRowNum = rowNum
        sheet._addRow(SheetRow { it.sheet = sheet; it.index = rowNum })
        continue
      }

      // find </row> closing tag
      re := xml.index("</row>", gt)
      if (re == null) break
      re += 6

      // wrap row in worksheet tag to preserve namespace declarations
      rowXml := wsTag + xml[rs..<re] + wsSuffix
      pos = re
      xr := XParser(rowXml.in).parseDoc.root.elems.first
      readRow(xr, sheet, sst, lastRowNum) |newLast| { lastRowNum = newLast }
    }

    // trim trailing empty rows
    sheet.trim
  }

  ** Process a parsed row element into sheet data.
  private static Void readRow(XElem xr, Sheet sheet, Int:Str sst,
                              Int lastRowNum, |Int| updateLast)
  {
    rowNum := xr.attr("r", false)?.val?.toInt ?: (lastRowNum + 1)
    if (rowNum > maxRows) throw ParseErr("Row $rowNum exceeds max ($maxRows)")

    // backfill sparse/missing rows
    cur := lastRowNum
    while (rowNum > cur + 1)
    {
      cur++
      sheet._addRow(SheetRow { it.sheet = sheet; it.index = cur })
    }
    updateLast(rowNum)

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
      if (cix >= maxCols) throw ParseErr("Column $cix exceeds max ($maxCols)")
      miss := cix - lastcix
      while (miss-- > 1) row._addCell(SheetCell { it.val="" })
      lastcix = cix

      // read value based on type
      switch (type)
      {
        case "s":
          // shared string
          sid  := xc.elems.first.text.val.toInt
          val  := sst[sid] ?: ""
          row._addCell(SheetCell { it.val=val })

        case "b":
          // boolean: <v>0</v> or <v>1</v>
          bval := xc.elems.first?.text?.val
          row._addCell(SheetCell { it.val = bval == "1" ? "true" : "false" })

        case "inlineStr":
          // inline string: <is><t>value</t></is>
          val  := xc.elems.first?.elems?.first?.text?.val ?: ""
          row._addCell(SheetCell { it.val=val })

        default:
          val  := xc.elems.first?.text?.val ?: ""
          row._addCell(SheetCell { it.val=val })
      }
    }

    sheet._addRow(row)
  }
}