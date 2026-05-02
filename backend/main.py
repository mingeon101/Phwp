import zipfile, io, re, traceback, logging
from pathlib import Path
from urllib.parse import quote
import pdfplumber
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF to HWPX Converter")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ── 실제 한컴 HWPX 파일에서 추출한 고정 템플릿 데이터 ──
TEMPLATE_PATH = "/opt/render/project/src/backend/template.hwpx"

def load_template():
    with zipfile.ZipFile(TEMPLATE_PATH) as z:
        files = {name: z.read(name) for name in z.namelist()}
    section0 = files["Contents/section0.xml"].decode("utf-8")
    sec_tag_start = section0.find('<hs:sec')
    sec_root_end  = section0.find('>', sec_tag_start) + 1
    sec_root      = section0[sec_tag_start:sec_root_end]
    first_p_end   = section0.find('</hp:p>') + len('</hp:p>')
    first_para    = section0[sec_root_end:first_p_end]
    first_para_no_text = re.sub(r'<hp:t[^>]*>.*?</hp:t>', '', first_para, flags=re.DOTALL)
    return files, sec_root, first_para_no_text

try:
    TMPL_FILES, SEC_ROOT, FIRST_PARA = load_template()
    logger.info("템플릿 로드 완료")
except Exception as e:
    logger.error(f"템플릿 로드 실패: {e}")
    TMPL_FILES, SEC_ROOT, FIRST_PARA = None, None, None


def xml_escape(s):
    return (s.replace("&","&amp;").replace("<","&lt;")
             .replace(">","&gt;").replace('"',"&quot;"))

def make_para(pid, text):
    safe = xml_escape(text.strip())
    if safe:
        return (f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
                f'pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0"><hp:t>{safe}</hp:t></hp:run>'
                f'</hp:p>')
    return (f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
            f'pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0"/></hp:p>')

def make_hwpx(paragraphs):
    all_lines = []
    for text in paragraphs:
        for line in text.replace("\r\n","\n").replace("\r","\n").split("\n"):
            all_lines.append(line.strip())

    text_paras = "".join(make_para(i, l) for i, l in enumerate(all_lines))

    section0_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        + SEC_ROOT
        + FIRST_PARA
        + text_paras
        + '</hs:sec>'
    ).encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        # mimetype 비압축
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        z.writestr(zi, TMPL_FILES["mimetype"])
        # version.xml 비압축
        zi2 = zipfile.ZipInfo("version.xml")
        zi2.compress_type = zipfile.ZIP_STORED
        z.writestr(zi2, TMPL_FILES["version.xml"])
        # 나머지 실제 파일 그대로
        for name in ["settings.xml",
                     "META-INF/container.xml",
                     "META-INF/manifest.xml",
                     "META-INF/container.rdf",
                     "Contents/content.hpf",
                     "Contents/header.xml"]:
            z.writestr(name, TMPL_FILES[name])
        # section0만 교체
        z.writestr("Contents/section0.xml", section0_xml)
    return buf.getvalue()


@app.get("/")
async def root():
    return {"status": "ok", "template": TMPL_FILES is not None}

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if TMPL_FILES is None:
        raise HTTPException(500, "템플릿 파일이 없습니다. backend/template.hwpx 를 확인하세요.")
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

