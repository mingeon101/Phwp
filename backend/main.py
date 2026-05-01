"""
PDF → HWPX 변환 백엔드
HWPX = ZIP + XML (OWPML) 구조 — 한컴에서 100% 열림
"""
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from urllib.parse import quote
import pdfplumber, zipfile, io, traceback, logging, textwrap
from pathlib import Path

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF to HWPX Converter API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


def make_hwpx(paragraphs: list[str]) -> bytes:
    """HWPX (ZIP+XML/OWPML) 파일 생성"""
    buf = io.BytesIO()

    # 본문 XML 문단 생성
    para_xml = ""
    for text in paragraphs:
        escaped = (text.replace("&", "&amp;")
                       .replace("<", "&lt;")
                       .replace(">", "&gt;")
                       .replace('"', "&quot;"))
        for line in escaped.split("\n"):
            line = line.strip()
            para_xml += f'  <hp:p><hp:run><hp:rPr/><hp:t>{line}</hp:t></hp:run></hp:p>\n'

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:

        # mimetype (비압축)
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        z.writestr(zi, "application/hwp+zip")

        # META-INF/container.xml
        z.writestr("META-INF/container.xml", textwrap.dedent("""\
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <container xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
              <rootfiles>
                <rootfile full-path="Contents/content.hpf"
                          media-type="application/hwp+zip"/>
              </rootfiles>
            </container>"""))

        # Contents/content.hpf
        z.writestr("Contents/content.hpf", textwrap.dedent("""\
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <opf:package xmlns:opf="http://www.idpf.org/2007/opf"
                         unique-identifier="docId" version="2.0">
              <opf:metadata>
                <dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">변환 문서</dc:title>
              </opf:metadata>
              <opf:manifest>
                <opf:item id="header"   href="header.xml"   media-type="application/xml"/>
                <opf:item id="section0" href="section0.xml" media-type="application/xml"/>
                <opf:item id="settings" href="settings.xml" media-type="application/xml"/>
              </opf:manifest>
              <opf:spine>
                <opf:itemref idref="section0"/>
              </opf:spine>
            </opf:package>"""))

        # Contents/header.xml
        z.writestr("Contents/header.xml", textwrap.dedent("""\
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
                     xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
              <hh:refList>
                <hh:fontfaces>
                  <hh:fontface hh:lang="HANGUL">
                    <hh:font hh:name="함초롬바탕" hh:type="TTF"/>
                  </hh:fontface>
                  <hh:fontface hh:lang="LATIN">
                    <hh:font hh:name="함초롬바탕" hh:type="TTF"/>
                  </hh:fontface>
                </hh:fontfaces>
                <hh:charPrs>
                  <hh:charPr hh:id="0" hh:name="바탕글">
                    <hh:fontRef hh:lang="HANGUL" hh:face="함초롬바탕" hh:size="1000"/>
                    <hh:fontRef hh:lang="LATIN"  hh:face="함초롬바탕" hh:size="1000"/>
                  </hh:charPr>
                </hh:charPrs>
                <hh:paraPrs>
                  <hh:paraPr hh:id="0" hh:name="바탕글" hh:charPrIDRef="0">
                    <hh:paraMargin hh:left="0" hh:right="0"
                                   hh:prev="0" hh:next="0"/>
                    <hh:paraSpacing hh:lineSpacing="160"
                                    hh:lineSpacingType="PERCENT"/>
                  </hh:paraPr>
                </hh:paraPrs>
                <hh:styles>
                  <hh:style hh:type="para" hh:id="0" hh:name="바탕글"
                            hh:paraPrIDRef="0" hh:charPrIDRef="0"
                            hh:nextStyleIDRef="0" hh:langID="1042"/>
                </hh:styles>
              </hh:refList>
            </hh:head>"""))

        # Contents/section0.xml
        z.writestr("Contents/section0.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"\n'
            '        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">\n'
            '  <hs:secPr>\n'
            '    <hs:pgSz hp:w="59528" hp:h="84188"/>\n'
            '    <hs:pgMar hp:left="8504" hp:right="8504" hp:top="5668"\n'
            '              hp:bottom="4252" hp:header="4252" hp:footer="4252" hp:gutter="0"/>\n'
            '  </hs:secPr>\n'
            + para_xml +
            '</hs:sec>\n')

        # Contents/settings.xml
        z.writestr("Contents/settings.xml", textwrap.dedent("""\
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <hset:settings xmlns:hset="http://www.hancom.co.kr/hwpml/2011/settings">
            </hset:settings>"""))

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
