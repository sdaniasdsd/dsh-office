"""Generate a tiny, auditable OPC fixture (no office runtime or network needed)."""
import sys
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DOC = f'''<w:document xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:r="{R}">
<w:body>
  <w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
  </w:sectPr></w:pPr><w:r><w:t>第一节</w:t></w:r></w:p>
  <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>第二节</w:t>
    <w:br w:type="page"/><w:lastRenderedPageBreak/>
    <w:drawing><wp:anchor behindDoc="1" simplePos="0">
      <wp:positionH relativeFrom="margin"><wp:posOffset>-91440</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>182880</wp:posOffset></wp:positionV>
      <wp:extent cx="914400" cy="457200"/><wp:wrapSquare/>
      <wp:docPr id="1" name="图一" descr="测试图片"/>
      <a:graphic><a:graphicData><a:blip r:embed="rImg"/></a:graphicData></a:graphic>
    </wp:anchor></w:drawing>
    <w:drawing><wp:inline><wp:extent cx="100" cy="200"/><wp:docPr id="2" name="linked"/>
      <a:graphic><a:graphicData><a:blip r:link="rLink"/></a:graphicData></a:graphic>
    </wp:inline></w:drawing>
    <w:pict/>
  </w:r></w:p>
  <w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>跨页表</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
  <w:sectPr><w:type w:val="oddPage"/><w:pgSz w:w="15840" w:h="12240"/>
    <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/><w:cols w:num="2"/>
  </w:sectPr>
</w:body></w:document>'''

with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as archive:
    archive.writestr("word/document.xml", DOC)
    archive.writestr("[Content_Types].xml", '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Default Extension="png" ContentType="image/png"/>
    </Types>''')
    archive.writestr("word/_rels/document.xml.rels", f'''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rImg" Type="{R}/image" Target="media/image%201.png"/>
      <Relationship Id="rLink" Type="{R}/image" Target="https://example.invalid/image.png" TargetMode="External"/>
    </Relationships>''')
    archive.writestr("word/media/image 1.png", b"fixture image bytes; not rendered")
