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
    writePodFile(zip, `/docProps/core.xml`)
    writePodFile(zip, `/xl/theme/theme1.xml`)
    writePodFile(zip, `/xl/styles.xml`)

    // indexes
    writeWbXml(zip)
    writeWbRels(zip)
    writeAppXml(zip)
    writeContentTypes(zip)

    // sheets
    wb.eachSheet |s| { writeSheet(zip, s) }

    zip.close
  }

//////////////////////////////////////////////////////////////////////////
// Indexes
//////////////////////////////////////////////////////////////////////////

  ** Write workbook.xml index.
  private Void writeWbXml(Zip zip)
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
    wb.eachSheet |s|
    {
      xout.printLine("    <sheet name=\"${s.name}\" sheetId=\"${s.sheetId}\" r:id=\"${s.relId}\"/>")
    }
    xout.printLine(
     "  </sheets>
        <calcPr calcId=\"124519\" />
      </workbook>")
  }

  ** Write xl/_rels/workbook.xml.rels index.
  private Void writeWbRels(Zip zip)
  {
    maxRelId := 0
    xout := zip.writeNext(`/xl/_rels/workbook.xml.rels`)
    xout.printLine(
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">")

    wb.eachSheet |s|
    {
      // assume format: rId1
      maxRelId = maxRelId.max(s.relId[3..-1].toInt)
      xout.printLine("<Relationship Id=\"${s.relId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet${s.sheetId}.xml\"/>")
    }

    xout.printLine(
     "  <Relationship Id=\"rId${maxRelId+1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>
        <Relationship Id=\"rId${maxRelId+2}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>

      </Relationships>")
  }

  ** Write /docProps/app.xml index.
  private Void writeAppXml(Zip zip)
  {
    xout := zip.writeNext(`/docProps/app.xml`)
    xout.printLine(
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">
        <Application>Microsoft Excel</Application>
        <DocSecurity>0</DocSecurity>
        <ScaleCrop>false</ScaleCrop>
        <HeadingPairs>
          <vt:vector size=\"2\" baseType=\"variant\">
            <vt:variant>
              <vt:lpstr>Worksheets</vt:lpstr>
            </vt:variant>
            <vt:variant>
              <vt:i4>1</vt:i4>
            </vt:variant>
          </vt:vector>
        </HeadingPairs>
        <TitlesOfParts>
          <vt:vector size=\"${wb.sheets.size}\" baseType=\"lpstr\">")
    wb.eachSheet |s|
    {
      xout.printLine("      <vt:lpstr>${s.name}</vt:lpstr>")
    }
    xout.printLine(
     "    </vt:vector>
        </TitlesOfParts>
        <Company/>
        <LinksUpToDate>false</LinksUpToDate>
        <SharedDoc>false</SharedDoc>
        <HyperlinksChanged>false</HyperlinksChanged>
        <AppVersion>12.0000</AppVersion>
      </Properties>")
  }

  ** Write /[Content_Types].xml index.
  private Void writeContentTypes(Zip zip)
  {
    xout := zip.writeNext(`/[Content_Types].xml`)
    xout.printLine(
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
        <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
        <Default Extension=\"xml\" ContentType=\"application/xml\"/>
        <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>
        <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>
        <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>
        <Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>
        <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>")
    wb.eachSheet |s|
    {
      file := "sheet${s.sheetId}.xml"
      xout.printLine("<Override PartName=\"/xl/worksheets/${file}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>")
    }
    xout.printLine("</Types>")
  }

//////////////////////////////////////////////////////////////////////////
// Sheets
//////////////////////////////////////////////////////////////////////////

  ** Write sheet to zip file.
  private Void writeSheet(Zip zip, Sheet sheet)
  {
    file := "sheet${sheet.sheetId}.xml"
    sout := zip.writeNext(`/xl/worksheets/${file}`)
    sout.printLine(
     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
      <worksheet xmlns=\"${xmlns}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
        <dimension ref=\"A1:${sheet.lastRef}\"/>
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