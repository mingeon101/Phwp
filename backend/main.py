import zipfile, io, re, traceback, logging
from pathlib import Path
from urllib.parse import quote
import pdfplumber
from PIL import Image as PILImage
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF to HWPX Converter")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

TEMPLATE_PATH = "/opt/render/project/src/backend/template.hwpx"
TEXT_W = 42520  # 텍스트 영역 너비 (HWP 단위, A4 기준)

def load_template():
    with zipfile.ZipFile(TEMPLATE_PATH) as z:
        files = {name: z.read(name) for name in z.namelist()}
    s = files["Contents/section0.xml"].decode("utf-8")
    sec_start    = s.find('<hs:sec')
    sec_root_end = s.find('>', sec_start) + 1
    sec_root     = s[sec_start:sec_root_end]
    first_p_end  = s.find('</hp:p>') + len('</hp:p>')
    first_para   = s[sec_root_end:first_p_end]
    first_para_clean = re.sub(r'<hp:t[^>]*>.*?</hp:t>', '', first_para, flags=re.DOTALL)
    return files, sec_root, first_para_clean

try:
    TMPL_FILES, SEC_ROOT, FIRST_PARA = load_template()
    logger.info("템플릿 로드 완료")
except Exception as e:
    logger.error(f"템플릿 로드 실패: {e}")
    TMPL_FILES = SEC_ROOT = FIRST_PARA = None


def xe(s):
    return (s.replace("&","&amp;").replace("<","&lt;")
             .replace(">","&gt;").replace('"',"&quot;"))

def make_text_para(pid, text):
    safe = xe(text.strip())
    if safe:
        return (f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
                f'pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0"><hp:t>{safe}</hp:t></hp:run>'
                f'</hp:p>')
    return (f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
            f'pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0"/></hp:p>')

def make_picture_para(pid, bin_id, img_w, img_h):
    return (
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0" '
        f'pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">'
        f'<hp:picture treatAsChar="1" inlineAttr="1" numberingType="PICTURE" '
        f'textWrap="SQUARE" textFlow="LARGEST_ONLY" '
        f'textSideByEachOther="1" restrictInPage="0" preventOverlap="0" '
        f'zOrder="0" width="{img_w}" height="{img_h}" '
        f'horz="COLUMN" vert="PARA" horzRel="CONTENT" vertRel="CONTENT" '
        f'horzPos="0" vertPos="0">'
        f'<hp:sz width="{img_w}" height="{img_h}"/>'
        f'<hp:pos x="0" y="0"/>'
        f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:pictureInfo transparency="0" contrast="0" brightness="0" '
        f'effect="REAL_PIC" blendType="BLEND_NORMAL" alpha="255"/>'
        f'<hp:imageBorder type="NONE" width="0.1mm" color="#000000"/>'
        f'<hp:image binItemIDRef="{bin_id}" name="image{bin_id}.jpg"/>'
        f'<hp:clipRect left="0" top="0" right="{img_w}" bottom="{img_h}"/>'
        f'<hp:Effects/>'
        f'</hp:picture>'
        f'</hp:run>'
        f'</hp:p>'
    )

def make_hwpx(pages_text, pages_images):
    header = TMPL_FILES["Contents/header.xml"].decode("utf-8")

    # BinData 준비
    bin_datas = []
    for i, img_bytes in enumerate(pages_images):
        if img_bytes:
            bin_datas.append((i + 1, f"image{i+1}.jpg", img_bytes))

    # header.xml에 binDataList 추가
    if bin_datas:
        entries = "".join(
            f'<hh:binData id="{idx}" state="Embedding" format="jpg" '
            f'compress="DeflateWithSizePrefixedChunk" instream="1" filename="{fname}"/>'
            for idx, fname, _ in bin_datas
        )
        bin_list = f'<hh:binDataList count="{len(bin_datas)}">{entries}</hh:binDataList>'
        header = header.replace('</hh:refList>', bin_list + '</hh:refList>')

    # section0.xml 문단 조립
    paras = ""
    pid = 0
    for page_idx, (text, img_bytes) in enumerate(zip(pages_text, pages_images)):
        # 텍스트 문단
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            paras += make_text_para(pid, line.strip())
            pid += 1
        # 이미지 문단
        if img_bytes:
            img = PILImage.open(io.BytesIO(img_bytes))
            w, h = img.size
            img_h = int(TEXT_W * h / w)
            paras += make_picture_para(pid, page_idx + 1, TEXT_W, img_h)
            pid += 1

    section0 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        + SEC_ROOT + FIRST_PARA + paras + '</hs:sec>'
    ).encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        z.writestr(zi, TMPL_FILES["mimetype"])
        zi2 = zipfile.ZipInfo("version.xml")
        zi2.compress_type = zipfile.ZIP_STORED
        z.writestr(zi2, TMPL_FILES["version.xml"])
        for name in ["settings.xml", "META-INF/container.xml",
                     "META-INF/manifest.xml", "META-INF/container.rdf",
                     "Contents/content.hpf"]:
            z.writestr(name, TMPL_FILES[name])
        z.writestr("Contents/header.xml", header.encode("utf-8"))
        z.writestr("Contents/section0.xml", section0)
        for _, fname, jpeg in bin_datas:
            z.writestr(f"BinData/{fname}", jpeg)
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
        raise HTTPException(500, "템플릿 파일 없음. backend/template.hwpx 확인 필요.")
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "PDF 파일만 업로드 가능합니다.")
    content = await file.read()
    if len(content) > 50 * 1024 * 1024:
        raise HTTPException(400, "50MB 이하 파일만 지원합니다.")
    try:
        logger.info(f"변환 시작: {file.filename} ({len(content):,}B)")
        pages_text, pages_images = [], []
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                pages_text.append(text)
                # 이미지가 있는 페이지만 렌더링
                if page.images:
                    img = page.to_image(resolution=150)
                    ibuf = io.BytesIO()
                    img.original.convert('RGB').save(ibuf, format='JPEG', quality=85)
                    pages_images.append(ibuf.getvalue())
                    logger.info(f"  p{i+1}: {len(text)}자, 이미지 {len(page.images)}개")
                else:
                    pages_images.append(None)
                    logger.info(f"  p{i+1}: {len(text)}자")

        if not any(t.strip() for t in pages_text):
            pages_text = ["(이미지 기반 PDF — 텍스트 추출 불가)"]
            pages_images = [None]

        hwpx = make_hwpx(pages_text, pages_images)
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
