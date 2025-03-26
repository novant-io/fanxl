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
    writePodFile(zip, `/xl/workbook.xml`)
    writePodFile(zip, `/[Content_Types].xml`)

    // sheets
    wb.sheets.each |sheet|
    {
      sout := zip.writeNext(`/xl/worksheets/sheet1.xml`)
      sout.printLine(
       Str<|<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <dimension ref="A1:E10"/>
            <sheetViews>
              <sheetView tabSelected="1" workbookViewId="0"/>
            </sheetViews>
            <sheetFormatPr defaultRowHeight="15"/>
            <sheetData>|>)

      // TODO: rows + shared_strings

      sout.printLine(
       Str<|</sheetData>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>|>)
      sout.close
    }

    zip.close
  }

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

  private Workbook wb
}