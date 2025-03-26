#! /usr/bin/env fan

using build

class Build : build::BuildPod
{
  new make()
  {
    podName = "fanxl"
    summary = "Fantom API for parsing Excel XLSX files"
    version = Version("0.1")
    meta = [
      "org.name":     "Novant",
      "org.uri":      "https://novant.io/",
      "license.name": "MIT",
      "vcs.name":     "Git",
      "vcs.uri":      "https://github.com/novant-io/fanxl"
    ]
    depends = ["sys 1.0", "util 1.0", "xml 1.0"]
    srcDirs = [
      `fan/`,
      `fan/csv/`,
      // `fan/num/`,
      `fan/xls/`,
      `test/`
    ]
    resDirs = [
      `res/template-xls/`,
      `res/template-xls/_rels/`,
      `res/template-xls/docProps/`,
      `res/template-xls/xl/`,
      `res/template-xls/xl/_rels/`,
      `res/template-xls/xl/theme/`,
      `res/template-xls/xl/worksheets/`,
    ]
    docApi  = true
    docSrc  = true
  }
}
