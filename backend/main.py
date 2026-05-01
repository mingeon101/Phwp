"""
PDF → HWPX 변환 백엔드
실제 한컴 HWPX 파일 구조 기반 (tech.hancom.com 공식 문서 참고)
"""
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from urllib.parse import quote
import pdfplumber, zipfile, io, traceback, logging
from pathlib import Path

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF to HWPX Converter API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


def make_hwpx(paragraphs: list) -> bytes:
    buf = io.BytesIO()

    # 본문 문단 생성
    paras = ""
    for text in paragraphs:
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            line = line.strip()
            if not line:
                line = " "
            safe = (line.replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                        .replace('"', "&quot;")
                        .replace("'", "&apos;"))
            paras += (
                '  <hp:p>\n'
                '    <hp:pPr><hp:pStyle hp:val="Normal"/></hp:pPr>\n'
                '    <hp:run><hp:rPr/>'
                f'<hp:t xml:space="preserve">{safe}</hp:t>'
                '</hp:run>\n'
                '  </hp:p>\n'
            )

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:

        # 1. mimetype — 반드시 첫 번째, 비압축
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        z.writestr(zi, "application/hwp+zip")

        # 2. META-INF/container.xml
        z.writestr("META-INF/container.xml",
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container">\n'
            '  <rootfiles>\n'
            '    <rootfile full-path="Contents/content.hpf"'
            ' media-type="application/hwp+zip"/>\n'
            '  </rootfiles>\n'
            '</container>')

        # 3. META-INF/manifest.xml
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

        # 4. version.xml (루트)
        z.writestr("version.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<ver:versionsupport\n'
            '    xmlns:ver="http://www.hancom.co.kr/hwpml/2011/versionsupport"\n'
            '    ver:VersionHWPX="1.3.0.0"\n'
            '    ver:VersionOWPML="1.1.0.0">\n'
            '</ver:versionsupport>')

        # 5. settings.xml (루트)
        z.writestr("settings.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<ha:HWPApplicationSetting\n'
            '    xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"\n'
            '    xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">\n'
            '  <ha:CaretPosition ha:listIDRef="0" ha:paraIDRef="0" ha:pos="0"/>\n'
            '</ha:HWPApplicationSetting>')

        # 6. Contents/content.hpf
        z.writestr("Contents/content.hpf",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<opf:package xmlns:opf="http://www.idpf.org/2007/opf"\n'
            '             unique-identifier="hwpx-document-id" version="2.0">\n'
            '  <opf:metadata>\n'
            '    <dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">변환 문서</dc:title>\n'
            '  </opf:metadata>\n'
            '  <opf:manifest>\n'
            '    <opf:item id="header"   href="header.xml"   media-type="application/xml"/>\n'
            '    <opf:item id="section0" href="section0.xml" media-type="application/xml"/>\n'
            '  </opf:manifest>\n'
            '  <opf:spine>\n'
            '    <opf:itemref idref="section0"/>\n'
            '  </opf:spine>\n'
            '</opf:package>')

        # 7. Contents/header.xml
        z.writestr("Contents/header.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<hh:head\n'
            '    xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"\n'
            '    xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"\n'
            '    xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">\n'
            '  <hh:beginNum hh:page="1" hh:footnote="1" hh:endnote="1"\n'
            '               hh:pic="1" hh:tbl="1" hh:equation="1"/>\n'
            '  <hh:refList>\n'
            '    <hh:fontfaces>\n'
            '      <hh:fontface hh:lang="HANGUL">\n'
            '        <hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/>\n'
            '      </hh:fontface>\n'
            '      <hh:fontface hh:lang="LATIN">\n'
            '        <hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/>\n'
            '      </hh:fontface>\n'
            '      <hh:fontface hh:lang="HANJA">\n'
            '        <hh:font hh:name="함초롬바탕" hh:type="TTF" hh:isEmbedded="0"/>\n'
            '      </hh:fontface>\n'
            '    </hh:fontfaces>\n'
            '    <hh:charPrs>\n'
            '      <hh:charPr hh:id="0" hh:name="바탕글" hh:height="1000"\n'
            '                 hh:textColor="0" hh:shadeColor="16777215" hh:useFontSpace="0"\n'
            '                 hh:useKerning="0" hh:symMark="NONE" hh:borderFillIDRef="0">\n'
            '        <hh:fontRef hh:lang="HANGUL" hh:face="함초롬바탕"/>\n'
            '        <hh:fontRef hh:lang="LATIN"  hh:face="함초롬바탕"/>\n'
            '        <hh:fontRef hh:lang="HANJA"  hh:face="함초롬바탕"/>\n'
            '        <hh:ratio hh:stretch="100" hh:lang="HANGUL"/>\n'
            '        <hh:ratio hh:stretch="100" hh:lang="LATIN"/>\n'
            '        <hh:spacing hh:letterSpacing="0" hh:lang="HANGUL"/>\n'
            '        <hh:spacing hh:letterSpacing="0" hh:lang="LATIN"/>\n'
            '        <hh:relSz hh:size="100" hh:lang="HANGUL"/>\n'
            '        <hh:relSz hh:size="100" hh:lang="LATIN"/>\n'
            '        <hh:offset hh:pos="0" hh:lang="HANGUL"/>\n'
            '        <hh:offset hh:pos="0" hh:lang="LATIN"/>\n'
            '      </hh:charPr>\n'
            '    </hh:charPrs>\n'
            '    <hh:tabPrs>\n'
            '      <hh:tabPr hh:id="0" hh:autoTabLeft="1" hh:autoTabRight="0"/>\n'
            '    </hh:tabPrs>\n'
            '    <hh:numberings/>\n'
            '    <hh:bullets/>\n'
            '    <hh:paraPrs>\n'
            '      <hh:paraPr hh:id="0" hh:name="바탕글" hh:tabPrIDRef="0"\n'
            '                 hh:condense="0" hh:fontLineHeight="0" hh:snapToGrid="1"\n'
            '                 hh:suppressLineNumbers="0" hh:checked="0">\n'
            '        <hh:align hh:horizontal="BOTH" hh:vertical="BASELINE"\n'
            '                  hh:useFontInfo="0" hh:textAlign="0"/>\n'
            '        <hh:heading hh:type="NONE" hh:idRef="0" hh:level="0"/>\n'
            '        <hh:breakSetting hh:breakLatinWord="KEEP_WORD" hh:breakNonLatinWord="KEEP_WORD"\n'
            '                         hh:widowOrphan="0" hh:keepWithNext="0"\n'
            '                         hh:keepLines="0" hh:pageBreakBefore="0" hh:columnBreakBefore="0"/>\n'
            '        <hh:margin hh:left="0" hh:right="0" hh:prev="0" hh:next="0"\n'
            '                   hh:indent="0" hh:charIndent="0"/>\n'
            '        <hh:lineSpacing hh:type="PERCENT" hh:value="160"/>\n'
            '        <hh:border hh:borderFillIDRef="0" hh:offsetLeft="0" hh:offsetRight="0"\n'
            '                   hh:offsetTop="0" hh:offsetBottom="0" hh:connect="0"\n'
            '                   hh:ignoreMargin="0"/>\n'
            '      </hh:paraPr>\n'
            '    </hh:paraPrs>\n'
            '    <hh:styles>\n'
            '      <hh:style hh:type="PARA" hh:id="0" hh:name="Normal"\n'
            '                hh:engName="Normal" hh:paraPrIDRef="0" hh:charPrIDRef="0"\n'
            '                hh:nextStyleIDRef="0" hh:langID="1042" hh:lockForm="0"/>\n'
            '    </hh:styles>\n'
            '    <hh:borderFills>\n'
            '      <hh:borderFill hh:id="0" hh:threeD="0" hh:shadow="0"\n'
            '                     hh:centerLine="0" hh:breakCellSeparateLine="0">\n'
            '        <hh:slash hh:type="NONE" hh:crooked="0" hh:isCounter="0"/>\n'
            '        <hh:backSlash hh:type="NONE" hh:crooked="0" hh:isCounter="0"/>\n'
            '        <hh:leftBorder hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:rightBorder hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:topBorder hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:bottomBorder hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:diagonal hh:type="NONE" hh:width="0.1mm" hh:color="0"/>\n'
            '        <hh:fillBrush>\n'
            '          <hh:noFill/>\n'
            '        </hh:fillBrush>\n'
            '      </hh:borderFill>\n'
            '    </hh:borderFills>\n'
            '    <hh:fillBrushes/>\n'
            '  </hh:refList>\n'
            '  <hh:compatibleDocument hh:targetProgram="HWP201X"/>\n'
            '  <hh:docOption>\n'
            '    <hh:linkinfo hh:path="" hh:pageInherit="1" hh:footnoteInherit="0"/>\n'
            '  </hh:docOption>\n'
            '  <hh:trackchageConfig hh:flags="0"/>\n'
            '</hh:head>')

        # 8. Contents/section0.xml
        z.writestr("Contents/section0.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<hs:sec\n'
            '    xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"\n'
            '    xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"\n'
            '    xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">\n'
            '  <hs:secPr hc:id="0" hs:textDirection="HORIZONTAL" hs:spaceColumns="1134"\n'
            '            hs:tabStop="8000" hs:outlineShapeIDRef="0"\n'
            '            hs:masterPageIDRef="0" hs:hideHeader="0" hs:hideFooter="0"\n'
            '            hs:hideMasterPage="0" hs:hidePageNumPos="0"\n'
            '            hs:hideBorderFill="0">\n'
            '    <hs:pgSz hp:w="59528" hp:h="84188" hs:orientation="PORTRAIT"/>\n'
            '    <hs:pgMar hp:left="8504" hp:right="8504" hp:top="5668"\n'
            '              hp:bottom="4252" hs:header="4252" hs:footer="4252" hs:gutter="0"/>\n'
            '    <hs:pageBorderFill hs:type="PAPER" hs:borderFillIDRef="0"\n'
            '                       hs:textOffsetLeft="1417" hs:textOffsetRight="1417"\n'
            '                       hs:textOffsetTop="1417" hs:textOffsetBottom="1417"\n'
            '                       hs:headerInside="0" hs:footerInside="0"/>\n'
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
        cd = f"attachment; filename*=UTF-8''{quote(stem + '.hwpx', safe='')}"
        return StreamingResponse(
            io.BytesIO(hwpx),
            media_type="application/hwp+zip",
            headers={
                "Content-Disposition": cd,
                "X-Pages": str(total),
                "Access-Control-Expose-Headers": "X-Pages,Content-Disposition",
            }
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(traceback.format_exc())
        raise HTTPException(500, f"변환 오류: {e}")

