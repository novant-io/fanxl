//
// Copyright (c) 2025, Novant LLC
// Licensed under the MIT License
//
// History:
//   20 Aug 2025  Andy Frank  Creation
//

*************************************************************************
** XmlWriter
*************************************************************************

** XmlWriter
internal class XmlWriter
{
  ** Create a new writer for given output stream.
  new make(OutStream out) { this.out = out }

  ** Write <?xml> declaration
  This decl(Bool standalone := true)
  {
    sa := standalone ? "yes" : "no"
    out.print("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"${sa}\"?>")
    return this
  }

  ** Start an element with given tag name and attrs.
  This elem(Str tag, [Str:Obj]? attrs := null)
  {
    out.writeChar('<').writeXml(tag)
    if (attrs != null)
    {
      attrs.each |v,k|
      {
        out.writeChar(' ')
          .writeXml(k)
          .writeChar('=')
          .writeChar('\"')
          .writeXml(v, OutStream.xmlEscQuotes.or(OutStream.xmlEscNewlines))
          .writeChar('\"')
      }
    }
    out.writeChar('>')
    return this
  }

  ** End a element with given tag.
  This elemEnd(Str tag)
  {
    out.writeChar('<').writeChar('/').writeXml(tag).writeChar('>')
    return this
  }

  private OutStream out
}