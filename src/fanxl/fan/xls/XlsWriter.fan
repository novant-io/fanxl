//
// Copyright (c) 2025, Novant LLC
// Licensed under the MIT License
//
// History:
//   26 Mar 2025  Andy Frank  Creation
//

using util
using xml

**
** XlsWriter writes a Workbook to XLSX file format.
**
internal class XlsWriter
{

//////////////////////////////////////////////////////////////////////////
// Construction
//////////////////////////////////////////////////////////////////////////

  ** Construct a new writer for given workbook.
  new make(Workbook wb)
  {
    this.wb = wb
  }

//////////////////////////////////////////////////////////////////////////
// Write
//////////////////////////////////////////////////////////////////////////

  ** Write workbook to given output stream.
  Void write(OutStream out)
  {
    // init zip
    zip := Zip.write(out)

    // doc metadata
    writePodFile(zip, `/_rels/.rels`)
    writePodFile(zip, `/docProps/app.xml`)
    writePodFile(zip, `/docProps/core.xml`)
    writePodFile(zip, `/xl/_rels/workbook.xml.rels`)
    writePodFile(zip, `/xl/theme/theme1.xml`)
    writePodFile(zip, `/xl/styles.xml`)
    writePodFile(zip, `/[Content_Types].xml`)

    // content
    writeIndex(zip)
    wb.sheets.each |s| { writeSheet(zip, s) }

    zip.close
  }

  ** Write workbook.xml index.
  private Void writeIndex(Zip zip)
  {
    xout := zip.writeNext(`/xl/workbook.xml`)
    xout.printLine(
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
        <fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/>
        <workbookPr defaultThemeVersion=\"124226\"/>
        <bookViews>
          <workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"16095\" windowHeight=\"9660\"/>
        </bookViews>
        <sheets>")
    wb.sheets.each |s|
    {
      xout.printLine("    <sheet name=\"${s.name}\" sheetId=\"${s.sheetId}\" r:id=\"${s.relId}\"/>")
    }
    xout.printLine(
     "  </sheets>
        <calcPr calcId=\"124519\" />
      </workbook>")
  }

  ** Write sheet to zip file.
  private Void writeSheet(Zip zip, Sheet sheet)
  {
    file := "sheet${sheet.sheetId}.xml"
    sout := zip.writeNext(`/xl/worksheets/${file}`)
    sout.printLine(
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <worksheet xmlns=\"${xmlns}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
        <dimension ref=\"A1:C1\"/>
        <sheetViews>
          <sheetView tabSelected=\"1\" workbookViewId=\"0\"/>
        </sheetViews>
        <sheetFormatPr defaultRowHeight=\"15\"/>
        <sheetData>")

    sheet.rows.each |row|
    {
      sout.printLine("<row r=\"${row.index}\" spans=\"1:3\">")
      3.times |i|
      {
        col  := ('A' + i).toChar
        sout.printLine("<c r=\"${col}${row.index}\" t=\"inlineStr\">")
        sout.printLine("<is><t>foo-${col}</t></is>")
        sout.printLine("</c>")
      }

      sout.printLine("</row>")
    }

    sout.printLine(
     Str<|</sheetData>
          <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
          </worksheet>|>)
    sout.close
  }

//////////////////////////////////////////////////////////////////////////
// Support
//////////////////////////////////////////////////////////////////////////

  ** Pipe a pod file to given zip file into zip.
  private Void writePodFile(Zip zip, Uri uri)
  {
    in := typeof.pod.file(`/res/template-xls${uri}`).in
    out := zip.writeNext(uri)
    in.pipe(out)
    out.close
    in.close
  }

//////////////////////////////////////////////////////////////////////////
// Fields
//////////////////////////////////////////////////////////////////////////

  private static const Str xmlns := "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  private Workbook wb
}