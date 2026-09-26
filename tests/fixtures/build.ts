import { strToU8, zipSync } from 'fflate';

/** 1x1 透明 PNG，作为图片素材。 */
export const PNG_1PX = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

const NS_PPTX =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const NS_CHART = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';

const NS_DOCX = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const NS_XLSX = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

type Files = Record<string, string | Uint8Array>;

function zip(files: Files): Uint8Array {
  const entries: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) {
    entries[k] = typeof v === 'string' ? strToU8(v) : v;
  }
  return zipSync(entries);
}

// ---------------------------------------------------------------------------
// PPTX
// ---------------------------------------------------------------------------

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>季度业务回顾</dc:title>
  <dc:creator>张伟</dc:creator>
  <cp:keywords>AI,RAG,业务</cp:keywords>
  <dcterms:created>2026-09-01T08:00:00Z</dcterms:created>
</cp:coreProperties>`;

const SLIDE1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS_PPTX}>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="标题 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:pPr lvl="0"/><a:r><a:rPr lang="zh-CN" sz="3200" b="1"/><a:t>季度业务回顾</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="内容占位符 2"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="5400000" cy="3200400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r><a:rPr lang="zh-CN" sz="1800" b="1"/><a:t>营收同比增长</a:t></a:r>
            <a:r><a:rPr lang="zh-CN" sz="1800"/><a:t> 35%</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="0"/>
            <a:r><a:rPr lang="zh-CN" sz="1800"/><a:t>AI助手接入3个业务线</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="1"/>
            <a:r><a:rPr lang="zh-CN" sz="1600"/><a:t>自动摘要准确率92%</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr lvl="0"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="zh-CN" sz="1400"/><a:t>数据截至2026年9月</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="4" name="图片 3" descr="架构示意图"/>
          <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="6096000" y="1600200"/><a:ext cx="2000000" cy="1500000"/></a:xfrm></p:spPr>
      </p:pic>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="5" name="图表 4"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="457200" y="5000000"/><a:ext cx="4000000" cy="1200000"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId5"/></a:graphicData></a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="6" name="表格 5"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="4600000" y="5000000"/><a:ext cx="3000000" cy="1200000"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1"/>
            <a:tblGrid><a:gridCol w="1500000"/><a:gridCol w="1500000"/></a:tblGrid>
            <a:tr><a:tc><a:txBody><a:p><a:r><a:t>业务线</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>营收</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
            <a:tr><a:tc><a:txBody><a:p><a:r><a:t>零售</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>1200</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
          </a:tbl>
        </a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>`;

const SLIDE2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS_PPTX}>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="标题 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr lang="zh-CN" sz="3200" b="1"/><a:t>下季度计划</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="文本框 2"/>
          <p:cNvSpPr txBox="1"/><p:nvPr/>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="5400000" cy="2400000"/></a:xfrm></p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square"/><a:lstStyle/>
          <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>上线文档解析服务</a:t></a:r></a:p>
          <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>把 doc2md 接到RAG流程</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

const SLIDE_HIDDEN = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS_PPTX} show="0">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="标题 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="3200"/><a:t>内部草稿页</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

const LAYOUT1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${NS_PPTX} type="obj">
  <p:cSld name="标题和内容">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="标题"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="内容"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle>
            <a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr>
            <a:lvl2pPr><a:buChar char="–"/></a:lvl2pPr>
          </a:lstStyle>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

const MASTER1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster ${NS_PPTX}>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="标题"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:buNone/></a:lvl1pPr></a:lstStyle></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="正文"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:buNone/></a:lvl1pPr></a:lstStyle></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sldMaster>`;

const NOTES1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes ${NS_PPTX}>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="幻灯片图像"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>‹#›</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="备注占位符"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>本页重点强调营收增长35%，</a:t></a:r></a:p>
          <a:p><a:r><a:t>并说明AI助手的落地情况。</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="页码"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{5F2C1B0A-0000-0000-0000-000000000000}" type="slidenum"><a:t>2</a:t></a:fld></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`;

const CHART1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${NS_CHART}>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p><a:r><a:t>季度营收趋势</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>营收(万元)</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>820</c:v></c:pt><c:pt idx="1"><c:v>910</c:v></c:pt><c:pt idx="2"><c:v>1107</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const PRESENTATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS_PPTX}>
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId3"/></p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
    <p:sldId id="258" r:id="rId4"/>
  </p:sldIdLst>
</p:presentation>`;

const PRESENTATION_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
</Relationships>`;

const SLIDE1_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;

const SLIDE2_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

const LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

const MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

/** 构造测试用 PPTX（含标题/要点/子级/表格/图表/图片/备注/隐藏页）。 */
export function buildPptx(): Uint8Array {
  return zip({
    '[Content_Types].xml': CONTENT_TYPES,
    '_rels/.rels': ROOT_RELS,
    'docProps/core.xml': CORE_XML,
    'ppt/presentation.xml': PRESENTATION,
    'ppt/_rels/presentation.xml.rels': PRESENTATION_RELS,
    'ppt/slides/slide1.xml': SLIDE1,
    'ppt/slides/slide2.xml': SLIDE2,
    'ppt/slides/slide3.xml': SLIDE_HIDDEN,
    'ppt/slides/_rels/slide1.xml.rels': SLIDE1_RELS,
    'ppt/slides/_rels/slide2.xml.rels': SLIDE2_RELS,
    'ppt/slideLayouts/slideLayout1.xml': LAYOUT1,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': LAYOUT_RELS,
    'ppt/slideMasters/slideMaster1.xml': MASTER1,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': MASTER_RELS,
    'ppt/notesSlides/notesSlide1.xml': NOTES1,
    'ppt/charts/chart1.xml': CHART1,
    'ppt/media/image1.png': PNG_1PX,
  });
}

// ---------------------------------------------------------------------------
// DOCX
// ---------------------------------------------------------------------------

const DOCX_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS_DOCX}>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/></w:style>
  <w:style w:type="character" w:styleId="CodeChar"><w:name w:val="HTML Code"/></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/></w:style>
</w:styles>`;

const DOCX_NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${NS_DOCX}>
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\uF0B7"/><w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const DOCX_FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${NS_DOCX}>
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1">
    <w:p>
      <w:r><w:footnoteRef/></w:r>
      <w:r><w:t>口径以财务系统为准。</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>`;

const DOCX_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS_DOCX} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>doc2md 项目说明</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>设计目标</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>把 PPTX、Word 文档转成</w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>干净的 Markdown</w:t></w:r>
      <w:r><w:t>，方便喂给 RAG 检索链路。</w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>中文排版不乱码。</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>安装命令：</w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>npm i doc2md</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>支持表格</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>支持嵌套列表</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>支持图片导出</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>第一阶段交付 CLI</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>第二阶段交付库</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Quote"/></w:pPr>
      <w:r><w:t>表格是 RAG 里最容易丢失的结构，必须保住。</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>参考</w:t></w:r>
      <w:hyperlink r:id="rId3">
        <w:r><w:rPr><w:color w:val="0563C1"/></w:rPr><w:t>官方文档</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t>与本页数据</w:t></w:r>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:footnoteReference w:id="1"/>
    </w:p>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>目录标题</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:t>后续说明</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:br w:type="page"/></w:r>
      <w:r><w:t>第二页内容</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr><w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:shd w:fill="DDDDDD"/></w:tcPr><w:p><w:r><w:t>模块</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>状态</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>pptx2md</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>已完成</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>docx2md 也在同一期</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="1905000" cy="1905000"/>
            <wp:docPr id="1" name="图片 1" descr="流程图"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:blipFill><a:blip r:embed="rId2"/></pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const DOCX_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/docs" TargetMode="External"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>`;

const DOCX_CORE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>doc2md 项目说明</dc:title>
  <dc:creator>李娜</dc:creator>
</cp:coreProperties>`;

/** 构造测试用 DOCX（标题/内联样式/列表/表格/图片/链接/脚注/域/分页）。 */
export function buildDocx(): Uint8Array {
  return zip({
    '[Content_Types].xml': CONTENT_TYPES,
    '_rels/.rels': ROOT_RELS.replace('ppt/presentation.xml', 'word/document.xml'),
    'docProps/core.xml': DOCX_CORE,
    'word/document.xml': DOCX_DOCUMENT,
    'word/_rels/document.xml.rels': DOCX_RELS,
    'word/styles.xml': DOCX_STYLES,
    'word/numbering.xml': DOCX_NUMBERING,
    'word/footnotes.xml': DOCX_FOOTNOTES,
    'word/media/image1.png': PNG_1PX,
  });
}

// ---------------------------------------------------------------------------
// XLSX
// ---------------------------------------------------------------------------

const XLSX_SHARED = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="9" uniqueCount="9">
  <si><t>产品</t></si>
  <si><t>区域</t></si>
  <si><t>销量</t></si>
  <si><t>单价</t></si>
  <si><t>华东</t></si>
  <si><t>华南</t></si>
  <si><r><t>doc2md</t></r><r><t>转换器</t></r></si>
  <si><t>工具</t></si>
  <si><t>汇总</t></si>
</sst>`;

const XLSX_SHEET1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS_XLSX}>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
      <c r="D1" t="s"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>6</v></c>
      <c r="B2" t="s"><v>4</v></c>
      <c r="C2"><v>120</v></c>
      <c r="D2"><v>199</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>7</v></c>
      <c r="B3" t="s"><v>5</v></c>
      <c r="C3"><v>80</v></c>
      <c r="D3"><v>59</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>8</v></c>
      <c r="B4" t="s"><v>4</v></c>
      <c r="C4" s="1"><v>0.35</v></c>
      <c r="D4" s="2"><v>46270</v></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>内联字符串</t></is></c>
      <c r="B5" t="b"><v>1</v></c>
    </row>
    <row r="6"></row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A8:C8"/></mergeCells>
</worksheet>`;

const XLSX_SHEET2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS_XLSX}>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>8</v></c><c r="B1"><v>1000</v></c></row>
  </sheetData>
</worksheet>`;

const XLSX_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="176" formatCode="yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/></numFmts>
  <cellXfs count="3">
    <xf numFmtId="0"/>
    <xf numFmtId="10"/>
    <xf numFmtId="176"/>
  </cellXfs>
</styleSheet>`;

const XLSX_WORKBOOK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook ${NS_XLSX}>
  <sheets>
    <sheet name="销售明细" sheetId="1" r:id="rId1"/>
    <sheet name="汇总" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`;

const XLSX_WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

/** 构造测试用 XLSX（多 sheet/共享字符串/数字格式/合并单元格）。 */
export function buildXlsx(): Uint8Array {
  return zip({
    '[Content_Types].xml': CONTENT_TYPES,
    '_rels/.rels': ROOT_RELS.replace('ppt/presentation.xml', 'xl/workbook.xml'),
    'docProps/core.xml': CORE_XML,
    'xl/workbook.xml': XLSX_WORKBOOK,
    'xl/_rels/workbook.xml.rels': XLSX_WORKBOOK_RELS,
    'xl/sharedStrings.xml': XLSX_SHARED,
    'xl/styles.xml': XLSX_STYLES,
    'xl/worksheets/sheet1.xml': XLSX_SHEET1,
    'xl/worksheets/sheet2.xml': XLSX_SHEET2,
  });
}
