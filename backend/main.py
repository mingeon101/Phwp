from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from urllib.parse import quote
import pdfplumber, zipfile, io, traceback, logging
from pathlib import Path

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF to HWPX Converter")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


def make_hwpx(paragraphs: list) -> bytes:
    buf = io.BytesIO()
    paras = ""
    pid = 0
    for text in paragraphs:
        for line in text.replace("\r\n","\n").replace("\r","\n").split("\n"):
            safe = (line.strip() or " ").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            paras += (
                f'  <hp:p hp:id="{pid}">\n'
                f'    <hp:pPr hp:paraPrIDRef="0" hp:tabPrIDRef="0" hp:styleIDRef="0"'
                f' hp:pageBreak="0" hp:columnBreak="0" hp:merged="0"/>\n'
                f'    <hp:run hp:charPrIDRef="0"><hp:rPr/>'
                f'<hp:t xml:space="preserve">{safe}</hp:t></hp:run>\n'
                f'  </hp:p>\n'
            )
            pid += 1

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        z.writestr(zi, "application/hwp+zip")

        z.writestr("META-INF/container.xml",
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container">\n'
            '  <rootfiles><rootfile full-path="Contents/content.hpf" media-type="application/hwp+zip"/></rootfiles>\n'
            '</container>')

        z.writestr("META-INF/manifest.xml",
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">\n'
            '  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/hwp+zip"/>\n'
            '  <manifest:file-entry manifest:full-path="Contents/content.hpf" manifest:media-type="application/xml"/>\n'
            '  <manifest:file-entry manifest:full-path="Contents/header.xml" manifest:media-type="application/xml"/>\n'
            '  <manifest:file-entry manifest:full-path="Contents/section0.xml" manifest:media-type="application/xml"/>\n'
            '  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="application/xml"/>\n'
            '  <manifest:file-entry manifest:full-path="version.xml" manifest:media-type="application/xml"/>\n'
            '</manifest:manifest>')

        z.writestr("META-INF/container.rdf",
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">\n'
            '  <rdf:Description rdf:about="Contents/content.hpf">\n'
            '    <rdf:type rdf:resource="http://www.idpf.org/2007/opf/components#Package"/>\n'
            '  </rdf:Description>\n'
            '</rdf:RDF>')

        z.writestr("version.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<ver:versionsupport xmlns:ver="http://www.hancom.co.kr/hwpml/2011/versionsupport">\n'
            '  <ver:hwpml ver:major="1" ver:minor="3" ver:micro="0" ver:build="0"/>\n'
            '</ver:versionsupport>')

        z.writestr("settings.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app">\n'
            '  <ha:CaretPosition ha:listIDRef="0" ha:paraIDRef="0" ha:pos="0"/>\n'
            '</ha:HWPApplicationSetting>')

        z.writestr("Contents/content.hpf",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<opf:package xmlns:opf="http://www.idpf.org/2007/opf" unique-identifier="hwpx-doc" version="2.0">\n'
            '  <opf:metadata><dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">문서</dc:title></opf:metadata>\n'
            '  <opf:manifest>\n'
            '    <opf:item id="header" href="header.xml" media-type="application/xml"/>\n'
            '    <opf:item id="section0" href="section0.xml" media-type="application/xml"/>\n'
            '  </opf:manifest>\n'
            '  <opf:spine><opf:itemref idref="section0"/></opf:spine>\n'
            '</opf:package>')

        z.writestr("Contents/header.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"\n'
            '         xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">\n'
            '  <hh:beginNum hh:page="1" hh:footnote="1" hh:endnote="1" hh:pic="1" hh:tbl="1" hh:equation="1"/>\n'
            '  <hh:refList>\n'
            '    <hh:fontfaces>\n'
            '      <hh:fontface hh:lang="HANGUL"><hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/></hh:fontface>\n'
            '      <hh:fontface hh:lang="LATIN"><hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/></hh:fontface>\n'
            '      <hh:fontface hh:lang="HANJA"><hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/></hh:fontface>\n'
            '      <hh:fontface hh:lang="JAPANESE"><hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/></hh:fontface>\n'
            '      <hh:fontface hh:lang="OTHER"><hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/></hh:fontface>\n'
            '      <hh:fontface hh:lang="SYMBOL"><hh:font hh:name="Symbol" hh:type="TTF" hh:isEmbedded="0"/></hh:fontface>\n'
            '      <hh:fontface hh:lang="USER"><hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/></hh:fontface>\n'
            '    </hh:fontfaces>\n'
            '    <hh:borderFills>\n'
            '      <hh:borderFill hh:id="0" hh:threeD="0" hh:shadow="0" hh:centerLine="0" hh:breakCellSeparateLine="0">\n'
            '        <hh:slash hh:type="NONE" hh:crooked="0" hh:isCounter="0"/>\n'
            '        <hh:backSlash hh:type="NONE" hh:crooked="0" hh:isCounter="0"/>\n'
            '        <hh:leftBorder hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:rightBorder hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:topBorder hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:bottomBorder hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:diagonal hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:fillBrush><hh:noFill/></hh:fillBrush>\n'
            '      </hh:borderFill>\n'
            '    </hh:borderFills>\n'
            '    <hh:charPrs>\n'
            '      <hh:charPr hh:id="0" hh:name="바탕글" hh:height="1000" hh:textColor="0"\n'
            '                 hh:shadeColor="16777215" hh:useFontSpace="0" hh:useKerning="0"\n'
            '                 hh:symMark="NONE" hh:borderFillIDRef="0">\n'
            '        <hh:fontRef hh:lang="HANGUL" hh:face="함초롬바탕"/>\n'
            '        <hh:fontRef hh:lang="LATIN" hh:face="함초롬바탕"/>\n'
            '        <hh:fontRef hh:lang="HANJA" hh:face="함초롬바탕"/>\n'
            '        <hh:fontRef hh:lang="JAPANESE" hh:face="함초롬바탕"/>\n'
            '        <hh:fontRef hh:lang="OTHER" hh:face="함초롬바탕"/>\n'
            '        <hh:fontRef hh:lang="SYMBOL" hh:face="Symbol"/>\n'
            '        <hh:fontRef hh:lang="USER" hh:face="함초롬바탕"/>\n'
            '        <hh:ratio hh:stretch="100" hh:lang="HANGUL"/><hh:ratio hh:stretch="100" hh:lang="LATIN"/>\n'
            '        <hh:ratio hh:stretch="100" hh:lang="HANJA"/><hh:ratio hh:stretch="100" hh:lang="JAPANESE"/>\n'
            '        <hh:ratio hh:stretch="100" hh:lang="OTHER"/><hh:ratio hh:stretch="100" hh:lang="SYMBOL"/>\n'
            '        <hh:ratio hh:stretch="100" hh:lang="USER"/>\n'
            '        <hh:spacing hh:letterSpacing="0" hh:lang="HANGUL"/><hh:spacing hh:letterSpacing="0" hh:lang="LATIN"/>\n'
            '        <hh:spacing hh:letterSpacing="0" hh:lang="HANJA"/><hh:spacing hh:letterSpacing="0" hh:lang="JAPANESE"/>\n'
            '        <hh:spacing hh:letterSpacing="0" hh:lang="OTHER"/><hh:spacing hh:letterSpacing="0" hh:lang="SYMBOL"/>\n'
            '        <hh:spacing hh:letterSpacing="0" hh:lang="USER"/>\n'
            '        <hh:relSz hh:size="100" hh:lang="HANGUL"/><hh:relSz hh:size="100" hh:lang="LATIN"/>\n'
            '        <hh:relSz hh:size="100" hh:lang="HANJA"/><hh:relSz hh:size="100" hh:lang="JAPANESE"/>\n'
            '        <hh:relSz hh:size="100" hh:lang="OTHER"/><hh:relSz hh:size="100" hh:lang="SYMBOL"/>\n'
            '        <hh:relSz hh:size="100" hh:lang="USER"/>\n'
            '        <hh:offset hh:pos="0" hh:lang="HANGUL"/><hh:offset hh:pos="0" hh:lang="LATIN"/>\n'
            '        <hh:offset hh:pos="0" hh:lang="HANJA"/><hh:offset hh:pos="0" hh:lang="JAPANESE"/>\n'
            '        <hh:offset hh:pos="0" hh:lang="OTHER"/><hh:offset hh:pos="0" hh:lang="SYMBOL"/>\n'
            '        <hh:offset hh:pos="0" hh:lang="USER"/>\n'
            '      </hh:charPr>\n'
            '    </hh:charPrs>\n'
            '    <hh:tabPrs><hh:tabPr hh:id="0" hh:autoTabLeft="1" hh:autoTabRight="0"/></hh:tabPrs>\n'
            '    <hh:numberings/>\n'
            '    <hh:bullets/>\n'
            '    <hh:paraPrs>\n'
            '      <hh:paraPr hh:id="0" hh:name="바탕글" hh:tabPrIDRef="0"\n'
            '                 hh:condense="0" hh:fontLineHeight="0" hh:snapToGrid="1"\n'
            '                 hh:suppressLineNumbers="0" hh:checked="0">\n'
            '        <hh:align hh:horizontal="BOTH" hh:vertical="BASELINE" hh:useFontInfo="0" hh:textAlign="0"/>\n'
            '        <hh:heading hh:type="NONE" hh:idRef="0" hh:level="0"/>\n'
            '        <hh:breakSetting hh:breakLatinWord="KEEP_WORD" hh:breakNonLatinWord="KEEP_WORD"\n'
            '                         hh:widowOrphan="0" hh:keepWithNext="0" hh:keepLines="0"\n'
            '                         hh:pageBreakBefore="0" hh:columnBreakBefore="0"/>\n'
            '        <hh:margin hh:left="0" hh:right="0" hh:prev="0" hh:next="0" hh:indent="0" hh:charIndent="0"/>\n'
            '        <hh:lineSpacing hh:type="PERCENT" hh:value="160"/>\n'
            '        <hh:border hh:borderFillIDRef="0" hh:offsetLeft="0" hh:offsetRight="0"\n'
            '                   hh:offsetTop="0" hh:offsetBottom="0" hh:connect="0" hh:ignoreMargin="0"/>\n'
            '      </hh:paraPr>\n'
            '    </hh:paraPrs>\n'
            '    <hh:styles>\n'
            '      <hh:style hh:type="PARA" hh:id="0" hh:name="바탕글" hh:engName="Normal"\n'
            '                hh:paraPrIDRef="0" hh:charPrIDRef="0" hh:nextStyleIDRef="0"\n'
            '                hh:langID="1042" hh:lockForm="0"/>\n'
            '    </hh:styles>\n'
            '  </hh:refList>\n'
            '  <hh:compatibleDocument hh:targetProgram="HWP201X"/>\n'
            '  <hh:docOption><hh:linkinfo hh:path="" hh:pageInherit="1" hh:footnoteInherit="0"/></hh:docOption>\n'
            '  <hh:trackchageConfig hh:flags="0"/>\n'
            '</hh:head>')

        z.writestr("Contents/section0.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"\n'
            '        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">\n'
            '  <hs:secPr>\n'
            '    <hs:pgSz hp:w="59528" hp:h="84188" hs:orientation="PORTRAIT"/>\n'
            '    <hs:pgMar hp:left="8504" hp:right="8504" hp:top="5668"\n'
            '              hp:bottom="4252" hs:header="4252" hs:footer="4252" hs:gutter="0"/>\n'
            '    <hs:pageBorderFill hs:type="PAPER" hs:borderFillIDRef="0"\n'
            '                       hs:textOffsetLeft="1417" hs:textOffsetRight="1417"\n'
            '                       hs:textOffsetTop="1417" hs:textOffsetBottom="1417"\n'
            '                       hs:headerInside="0" hs:footerInside="0"/>\n'
            '    <hs:columns hs:type="NEWSPAPER" hs:count="1" hs:spacing="1134"\n'
            '                hs:sameWidth="1" hs:oneLine="0" hs:direction="L2R" hs:balanceLastCol="0"/>\n'
            '    <hs:noHeader/>\n'
            '    <hs:noFooter/>\n'
            '    <hs:footnoteShape/>\n'
            '    <hs:endnoteShape/>\n'
            '  </hs:secPr>\n'
            + paras +
            '</hs:sec>\n')

    return buf.getvalue()


@app.get("/")
async def root():
    return {"status": "ok", "usage": "POST /convert"}

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "PDF 파일만 업로드 가능합니다.")
    content = await file.read()
    if len(content) > 50 * 1024 * 1024:
        raise HTTPException(400, "50MB 이하 파일만 지원합니다.")
    try:
        logger.info(f"변환 시작: {file.filename} ({len(content):,}B)")
        pages = []
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                t = page.extract_text() or ""
                pages.append(t)
                logger.info(f"  p{i+1}: {len(t)}자")
        if not any(t.strip() for t in pages):
            pages = ["(이미지 기반 PDF — 텍스트 추출 불가)"]
        hwpx = make_hwpx(pages)
        logger.info(f"HWPX 완성: {len(hwpx):,}B")
        stem = Path(file.filename).stem
        cd = f"attachment; filename*=UTF-8''{quote(stem+'.hwpx', safe='')}"
        return StreamingResponse(io.BytesIO(hwpx), media_type="application/hwp+zip",
            headers={"Content-Disposition": cd, "X-Pages": str(total),
                     "Access-Control-Expose-Headers": "X-Pages,Content-Disposition"})
    except HTTPException:
        raise
    except Exception as e:
        logger.error(traceback.format_exc())
        raise HTTPException(500, f"변환 오류: {e}")

