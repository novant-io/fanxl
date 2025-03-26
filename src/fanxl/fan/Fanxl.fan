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
    switch (file.ext.lower)
    {
      case "csv": return CsvReader.read(file)
      default:    return XlsReader.read(file)
    }
  }

  ** Write the given workbook in XLS format to output stream.
  static Void writeXls(Workbook wb, OutStream out)
  {
    XlsWriter(wb).write(out)
  }
}