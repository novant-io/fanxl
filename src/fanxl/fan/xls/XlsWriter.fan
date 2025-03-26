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
    zip  := Zip.write(out)
    writePodFile(zip, `/_rels/.rels`)
    writePodFile(zip, `/docProps/app.xml`)
    writePodFile(zip, `/docProps/core.xml`)
    writePodFile(zip, `/xl/_rels/workbook.xml.rels`)
    writePodFile(zip, `/xl/theme/theme1.xml`)
    writePodFile(zip, `/xl/worksheets/sheet1.xml`)
    writePodFile(zip, `/xl/styles.xml`)
    writePodFile(zip, `/xl/workbook.xml`)
    writePodFile(zip, `/[Content_Types].xml`)
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

  ** Write given file into zip.
  private Void writeNext(Zip zip, Uri uri, Str content)
  {
    out := zip.writeNext(uri)
    out.printLine(content)
    out.close
  }

//////////////////////////////////////////////////////////////////////////
// Fields
//////////////////////////////////////////////////////////////////////////


  private Workbook wb
}