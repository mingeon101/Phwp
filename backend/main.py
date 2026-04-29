"""
PDF → HWP 5.0 변환 백엔드
FastAPI + pdfplumber | Render.com 배포용
HWP 5.0 Revision 1.3 스펙 기반
"""

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from urllib.parse import quote
import pdfplumber
import struct
import zlib
import io
import traceback
import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF to HWP Converter API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

SECTOR_SIZE  = 512
FREESECT     = 0xFFFFFFFF
ENDOFCHAIN   = 0xFFFFFFFE
FATSECT      = 0xFFFFFFFD
NOSTREAM     = 0xFFFFFFFF
HWPTAG_BEGIN = 0x010


def pad512(data):
    r = len(data) % SECTOR_SIZE
    return data + b"\x00" * (SECTOR_SIZE - r if r else 0)


def make_dir_entry(name, typ, color, left, right, child, start, size, clsid=None):
    if clsid is None:
        clsid = b"\x00" * 16
    enc  = name.encode("utf-16-le") if name else b""
    nlen = len(enc) + 2 if name else 0
    nfld = (enc + b"\x00\x00")[:62].ljust(64, b"\x00")
    d  = nfld
    d += struct.pack("<H", nlen)
    d += struct.pack("<B", typ)
    d += struct.pack("<B", color)
    d += struct.pack("<I", left)
    d += struct.pack("<I", right)
    d += struct.pack("<I", child)
    d += clsid
    d += struct.pack("<I", 0)
    d += struct.pack("<Q", 0)
    d += struct.pack("<Q", 0)
    d += struct.pack("<I", start)
    d += struct.pack("<Q", size)
    assert len(d) == 128
    return d


def build_cfb(file_header, doc_info_z, section0_z):
    streams = [
        ("FileHeader",        file_header),
        ("DocInfo",           doc_info_z),
        ("BodyText/Section0", section0_z),
    ]
    sector_data = b""
    fat = []
    cur = 0
    locs = {}

    for name, data in streams:
        padded = pad512(data)
        n = len(padded) // SECTOR_SIZE
        locs[name] = (cur, len(data))
        sector_data += padded
        for j in range(n):
            fat.append(cur + j + 1 if j < n - 1 else ENDOFCHAIN)
        cur += n

    fh_s, fh_z = locs["FileHeader"]
    di_s, di_z = locs["DocInfo"]
    s0_s, s0_z = locs["BodyText/Section0"]

    entries = [
        make_dir_entry("Root Entry", 5, 1, NOSTREAM, NOSTREAM, 1,
                       ENDOFCHAIN, 0,
                       b"\x00\x20\x08\x02\x00\x00\x00\x00"
                       b"\xc0\x00\x00\x00\x00\x00\x00\x46"),
        make_dir_entry("FileHeader", 2, 1, NOSTREAM, 2, NOSTREAM, fh_s, fh_z),
        make_dir_entry("DocInfo",    2, 1, NOSTREAM, 3, NOSTREAM, di_s, di_z),
        make_dir_entry("BodyText",   1, 1, NOSTREAM, NOSTREAM, 4, ENDOFCHAIN, 0),
        make_dir_entry("Section0",   2, 1, NOSTREAM, NOSTREAM, NOSTREAM, s0_s, s0_z),
    ]
    while len(entries) % 4:
        entries.append(make_dir_entry("", 0, 1, NOSTREAM, NOSTREAM, NOSTREAM, ENDOFCHAIN, 0))

    dir_bytes = b"".join(entries)
    dir_start = cur
    n_dir = len(dir_bytes) // SECTOR_SIZE
    for j in range(n_dir):
        fat.append(dir_start + j + 1 if j < n_dir - 1 else ENDOFCHAIN)
    cur += n_dir

    fat_start = cur
    fat.append(FATSECT)
    while len(fat) % 128:
        fat.append(FREESECT)
    fat_bytes = struct.pack(f"<{len(fat)}I", *fat)
    n_fat = len(fat_bytes) // SECTOR_SIZE

    hdr  = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    hdr += b"\x00" * 16
    hdr += struct.pack("<H", 0x003E)
    hdr += struct.pack("<H", 0x0003)
    hdr += struct.pack("<H", 0xFFFE)
    hdr += struct.pack("<H", 9)
    hdr += struct.pack("<H", 6)
    hdr += b"\x00" * 6
    hdr += struct.pack("<I", 0)
    hdr += struct.pack("<I", n_fat)
    hdr += struct.pack("<I", dir_start)
    hdr += struct.pack("<I", 0)
    hdr += struct.pack("<I", 4096)
    hdr += struct.pack("<I", FREESECT)
    hdr += struct.pack("<I", 0)
    hdr += struct.pack("<I", FREESECT)
    hdr += struct.pack("<I", 0)
    hdr += struct.pack("<I", fat_start)
    hdr += struct.pack("<I", FREESECT) * 108
    assert len(hdr) == 512

    return hdr + sector_data + dir_bytes + fat_bytes


def hwp_record(tag_id, level, data):
    size = len(data)
    if size >= 0xFFF:
        hdr = (tag_id & 0x3FF) | ((level & 0x3FF) << 10) | (0xFFF << 20)
        return struct.pack("<II", hdr, size) + data
    hdr = (tag_id & 0x3FF) | ((level & 0x3FF) << 10) | ((size & 0xFFF) << 20)
    return struct.pack("<I", hdr) + data


def make_file_header():
    sig     = b"HWP Document File\x00" + b"\x00" * 14
    fh      = sig[:32]
    fh     += struct.pack("<I", 0x05000306)  # version 5.0.3.6
    fh     += struct.pack("<I", 0x00000001)  # attr1: bit0=압축
    fh     += struct.pack("<I", 0x00000000)
    fh     += struct.pack("<I", 0x00000000)
    fh     += struct.pack("<B", 0x00)
    fh     += b"\x00" * 207
    assert len(fh) == 256
    return fh


def make_doc_info():
    out = b""

    # HWPTAG_DOCUMENT_PROPERTIES
    dp  = struct.pack("<H", 1)
    dp += struct.pack("<HHHHHH", 1, 1, 1, 1, 1, 1)
    dp += struct.pack("<III", 0, 0, 0)
    out += hwp_record(HWPTAG_BEGIN + 0, 0, dp)

    # HWPTAG_ID_MAPPINGS
    im = struct.pack("<18i", 0,1,0,0,1,1,0,0,0,1,1,0,0,0,0,0,0,0)
    out += hwp_record(HWPTAG_BEGIN + 1, 0, im)

    # HWPTAG_FACE_NAME
    fname = "나눔고딕"
    fenc  = fname.encode("utf-16-le")
    face  = struct.pack("<B", 0)
    face += struct.pack("<H", len(fname)) + fenc
    face += struct.pack("<B", 0)
    face += struct.pack("<H", 0)
    face += b"\x00" * 10
    face += struct.pack("<H", 0)
    out += hwp_record(HWPTAG_BEGIN + 3, 1, face)

    # HWPTAG_CHAR_SHAPE 72 bytes
    cs  = struct.pack("<7H", *([0]*7))       # 14
    cs += struct.pack("<7B", *([100]*7))     # 7
    cs += struct.pack("<7b", *([0]*7))       # 7
    cs += struct.pack("<7B", *([100]*7))     # 7
    cs += struct.pack("<7b", *([0]*7))       # 7
    cs += struct.pack("<i",  1000)           # 4
    cs += struct.pack("<I",  0)              # 4
    cs += struct.pack("<bb", 0, 0)           # 2
    cs += struct.pack("<I",  0x000000)       # 4
    cs += struct.pack("<I",  0x000000)       # 4
    cs += struct.pack("<I",  0x808080)       # 4
    cs += struct.pack("<I",  0x000000)       # 4
    cs += struct.pack("<H",  0)              # 2
    cs += struct.pack("<H",  0)              # 2
    assert len(cs) == 72, f"cs={len(cs)}"
    out += hwp_record(HWPTAG_BEGIN + 5, 1, cs)

    # HWPTAG_PARA_SHAPE 54 bytes
    ps  = struct.pack("<I",    0)            # 4
    ps += struct.pack("<iiii", 0,0,0,0)      # 16
    ps += struct.pack("<ii",   0,0)          # 8
    ps += struct.pack("<i",    200)          # 4
    ps += struct.pack("<HHH",  0,0,0)        # 6
    ps += struct.pack("<hhhh", 0,0,0,0)      # 8
    ps += struct.pack("<II",   0,200)        # 8
    assert len(ps) == 54, f"ps={len(ps)}"
    out += hwp_record(HWPTAG_BEGIN + 9, 1, ps)

    # HWPTAG_STYLE
    sn   = "바탕글"
    senc = sn.encode("utf-16-le")
    st   = struct.pack("<H", len(sn)) + senc
    st  += struct.pack("<H", 0)
    st  += struct.pack("<B", 0)
    st  += struct.pack("<B", 0)
    st  += struct.pack("<h", 0)
    st  += struct.pack("<HH", 0, 0)
    out += hwp_record(HWPTAG_BEGIN + 10, 1, st)

    return out


def make_body(pages_text):
    out = b""
    for i, text in enumerate(pages_text):
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if text.strip() else [""]
        for line in lines:
            line_text = line + "\r"
            encoded   = line_text.encode("utf-16-le")
            nchars    = len(line_text)
            ph  = struct.pack("<I", nchars)
            ph += struct.pack("<I", 0)
            ph += struct.pack("<H", 0)
            ph += struct.pack("<B", 0)
            ph += struct.pack("<B", 0)
            ph += struct.pack("<H", 1)
            ph += struct.pack("<H", 0)
            ph += struct.pack("<H", 0)
            ph += struct.pack("<I", abs(hash(line_text)) & 0xFFFFFFFF)
            ph += struct.pack("<H", 0)
            out += hwp_record(HWPTAG_BEGIN + 50, 0, ph)
            out += hwp_record(HWPTAG_BEGIN + 51, 1, encoded)
            out += hwp_record(HWPTAG_BEGIN + 52, 1, struct.pack("<II", 0, 0))

        if i < len(pages_text) - 1:
            sep = "\r".encode("utf-16-le")
            ph  = struct.pack("<I", 1) + struct.pack("<I", 0) + struct.pack("<H", 0)
            ph += struct.pack("<B", 0) + struct.pack("<B", 0) + struct.pack("<H", 1)
            ph += struct.pack("<H", 0) + struct.pack("<H", 0)
            ph += struct.pack("<I", 0) + struct.pack("<H", 0)
            out += hwp_record(HWPTAG_BEGIN + 50, 0, ph)
            out += hwp_record(HWPTAG_BEGIN + 51, 1, sep)
            out += hwp_record(HWPTAG_BEGIN + 52, 1, struct.pack("<II", 0, 0))
    return out


@app.get("/")
async def root():
    return {"status": "ok", "service": "PDF to HWP Converter", "usage": "POST /convert"}

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
        logger.info(f"변환 시작: {file.filename} ({len(content):,} bytes)")
        pages_text = []
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                t = page.extract_text() or ""
                pages_text.append(t)
                logger.info(f"  페이지 {i+1}: {len(t)}자")

        if not any(t.strip() for t in pages_text):
            pages_text = ["(이미지 기반 PDF — 텍스트 추출 불가)"]

        doc_info = make_doc_info()
        body     = make_body(pages_text)
        hwp = build_cfb(
            file_header = make_file_header(),
            doc_info_z  = zlib.compress(doc_info, 6),
            section0_z  = zlib.compress(body,     6),
        )
        logger.info(f"HWP 완성: {len(hwp):,} bytes")

        stem    = Path(file.filename).stem
        encoded = quote(stem + ".hwp", safe="")
        cd      = f"attachment; filename*=UTF-8''{encoded}"
        return StreamingResponse(
            io.BytesIO(hwp),
            media_type="application/x-hwp",
            headers={
                "Content-Disposition": cd,
                "X-Pages": str(total),
                "Access-Control-Expose-Headers": "X-Pages,Content-Disposition",
            }
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"변환 오류:\n{traceback.format_exc()}")
        raise HTTPException(500, f"변환 오류: {str(e)}")
    style = struct.pack("<H", len(sname)) + senc
    style += struct.pack("<H", 0) + struct.pack("<B", 0) + struct.pack("<B", 0)
    style += struct.pack("<h", 0) + struct.pack("<HH", 0, 0)
    records += make_record(HWPTAG_BEGIN + 10, 1, style)
    return records

def text_to_para_records(text: str) -> bytes:
    records = b""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for line in lines:
        line_text = line + "\r"
        encoded = line_text.encode("utf-16-le")
        nchars = len(line_text)
        ph = struct.pack("<I", nchars)
        ph += struct.pack("<I", 0)
        ph += struct.pack("<H", 0)
        ph += struct.pack("<B", 0)
        ph += struct.pack("<B", 0)
        ph += struct.pack("<H", 1)
        ph += struct.pack("<H", 0)
        ph += struct.pack("<H", 0)
        ph += struct.pack("<I", abs(hash(line)) & 0xFFFFFFFF)
        ph += struct.pack("<H", 0)
        records += make_record(HWPTAG_BEGIN + 50, 0, ph)
        records += make_record(HWPTAG_BEGIN + 51, 1, encoded)
        records += make_record(HWPTAG_BEGIN + 52, 1, struct.pack("<II", 0, 0))
    return records

class CFBWriter:
    SECTOR_SIZE = 512
    FREESECT = 0xFFFFFFFF
    ENDOFCHAIN = 0xFFFFFFFE
    FATSECT = 0xFFFFFFFD
    NOSTREAM = 0xFFFFFFFF

    def __init__(self):
        self.streams: dict[str, bytes] = {}

    def add_stream(self, name: str, data: bytes):
        self.streams[name] = data

    def _pad(self, data: bytes) -> bytes:
        r = len(data) % self.SECTOR_SIZE
        return data + b"\x00" * (self.SECTOR_SIZE - r if r else 0)

    def _dir_entry(self, name, typ, color, left, right, child, start, size, clsid=None) -> bytes:
        if clsid is None:
            clsid = b"\x00" * 16
        enc = name.encode("utf-16-le") if name else b""
        name_len = len(enc) + 2 if name else 0
        name_field = (enc + b"\x00\x00")[:62].ljust(64, b"\x00")
        d = name_field
        d += struct.pack("<H", name_len)
        d += struct.pack("<BB", typ, color)
        d += struct.pack("<II", left, right)
        d += struct.pack("<I", child)
        d += clsid
        d += struct.pack("<I", 0)
        d += struct.pack("<QQ", 0, 0)
        d += struct.pack("<I", start)
        d += struct.pack("<Q", size)
        assert len(d) == 128
        return d

    def build(self) -> bytes:
        items = list(self.streams.items())
        sector_blobs = []
        fat = []
        cur = 0

        for name, data in items:
            padded = self._pad(data)
            n = len(padded) // self.SECTOR_SIZE
            sector_blobs.append(padded)
            for j in range(n):
                fat.append(cur + j + 1 if j < n - 1 else self.ENDOFCHAIN)
            cur += n

        # directory
        dir_start = cur
        entries = []
        # root
        entries.append(self._dir_entry("Root Entry", 5, 1,
            self.NOSTREAM, self.NOSTREAM, 1 if items else self.NOSTREAM,
            self.ENDOFCHAIN, 0,
            b"\x00\x20\x08\x02\x00\x00\x00\x00\xc0\x00\x00\x00\x00\x00\x00\x46"))
        # streams
        sec = 0
        for i, (name, data) in enumerate(items):
            n = len(self._pad(data)) // self.SECTOR_SIZE
            right = i + 2 if i + 1 < len(items) else self.NOSTREAM
            entries.append(self._dir_entry(name, 2, 1,
                self.NOSTREAM, right, self.NOSTREAM, sec, len(data)))
            sec += n

        # pad dir to sector boundary
        while len(entries) % 4:
            entries.append(self._dir_entry("", 0, 1,
                self.NOSTREAM, self.NOSTREAM, self.NOSTREAM, self.ENDOFCHAIN, 0))

        dir_data = b"".join(entries)
        n_dir = len(dir_data) // self.SECTOR_SIZE
        for j in range(n_dir):
            fat.append(dir_start + j + 1 if j < n_dir - 1 else self.ENDOFCHAIN)
        cur += n_dir

        # FAT sector
        fat_start = cur
        fat.append(self.FATSECT)
        fat_padded = fat + [self.FREESECT] * (128 - (len(fat) % 128 or 128))
        fat_data = self._pad(struct.pack(f"<{len(fat_padded)}I", *fat_padded))

        # Header
        hdr = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
        hdr += b"\x00" * 16          # CLSID
        hdr += struct.pack("<HH", 0x003E, 0x0003)  # minor/major ver
        hdr += struct.pack("<H", 0xFFFE)            # byte order
        hdr += struct.pack("<HH", 9, 6)             # sector/mini sector size
        hdr += b"\x00" * 6
        hdr += struct.pack("<I", 0)                 # num dir sectors
        hdr += struct.pack("<I", 1)                 # num FAT sectors
        hdr += struct.pack("<I", dir_start)         # dir start
        hdr += struct.pack("<I", 0)                 # transaction sig
        hdr += struct.pack("<I", 4096)              # mini stream cutoff
        hdr += struct.pack("<I", self.ENDOFCHAIN)   # mini FAT start
        hdr += struct.pack("<I", 0)                 # num mini FAT
        hdr += struct.pack("<I", self.ENDOFCHAIN)   # DIFAT start
        hdr += struct.pack("<I", 0)                 # num DIFAT
        hdr += struct.pack("<I", fat_start)         # first FAT sector
        hdr += struct.pack("<I", self.FREESECT) * 108
        assert len(hdr) == 512

        result = hdr
        for blob in sector_blobs:
            result += blob
        result += dir_data
        result += fat_data
        return result


@app.get("/")
async def root():
    return {"status": "ok", "message": "PDF to HWP Converter API", "usage": "POST /convert"}

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
        logger.info(f"변환 시작: {file.filename} ({len(content)} bytes)")

        # 1. PDF 텍스트 추출
        pages_text = []
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            total = len(pdf.pages)
            logger.info(f"PDF 페이지 수: {total}")
            for i, page in enumerate(pdf.pages):
                txt = page.extract_text() or ""
                pages_text.append(txt)
                logger.info(f"페이지 {i+1}: {len(txt)}자 추출")

        # 2. HWP 본문 레코드 생성
        body = b""
        for i, txt in enumerate(pages_text):
            if txt.strip():
                body += text_to_para_records(txt)
            if i < len(pages_text) - 1:
                body += text_to_para_records("")
        if not body:
            body = text_to_para_records("(변환된 텍스트 없음)")
        logger.info(f"본문 레코드: {len(body)} bytes")

        # 3. DocInfo 생성
        doc_info = make_doc_info()
        logger.info(f"DocInfo: {len(doc_info)} bytes")

        # 4. CFB 생성
        cfb = CFBWriter()
        cfb.add_stream("FileHeader", make_file_header())
        cfb.add_stream("DocInfo", zlib.compress(doc_info))
        cfb.add_stream("BodyText/Section0", zlib.compress(body))
        hwp = cfb.build()
        logger.info(f"HWP 생성 완료: {len(hwp)} bytes")

        stem = Path(file.filename).stem
        # RFC 5987: 한글 파일명 인코딩
        from urllib.parse import quote
        encoded = quote(stem + ".hwp", safe="")
        cd = f"attachment; filename*=UTF-8''{encoded}"
        return StreamingResponse(
            io.BytesIO(hwp),
            media_type="application/x-hwp",
            headers={
                "Content-Disposition": cd,
                "X-Pages": str(total),
                "Access-Control-Expose-Headers": "X-Pages,Content-Disposition",
            }
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"변환 오류:\n{traceback.format_exc()}")
        raise HTTPException(500, f"변환 오류: {str(e)}")
