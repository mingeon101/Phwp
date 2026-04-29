"""
PDF → HWP 5.0 변환 백엔드
FastAPI + pdfplumber
Render.com 배포용
"""

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
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
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)

# ──────────────────────────────────────────────
# HWP 5.0 바이너리 생성
# ──────────────────────────────────────────────

HWPTAG_BEGIN = 0x010

def make_record(tag_id: int, level: int, data: bytes) -> bytes:
    size = len(data)
    if size >= 0xFFF:
        header = (tag_id & 0x3FF) | ((level & 0x3FF) << 10) | (0xFFF << 20)
        return struct.pack("<II", header, size) + data
    header = (tag_id & 0x3FF) | ((level & 0x3FF) << 10) | ((size & 0xFFF) << 20)
    return struct.pack("<I", header) + data

def make_file_header() -> bytes:
    sig = b"HWP Document File\x00" + b"\x00" * 14
    version = struct.pack("<I", 0x05000306)
    attr1 = struct.pack("<I", 0x00000001)  # bit0: 압축
    attr2 = struct.pack("<I", 0x00000000)
    encrypt_ver = struct.pack("<I", 0x00000000)
    kogl = struct.pack("<B", 0x00)
    reserved = b"\x00" * 207
    header = sig[:32] + version + attr1 + attr2 + encrypt_ver + kogl + reserved
    assert len(header) == 256
    return header

def make_doc_info() -> bytes:
    records = b""
    # HWPTAG_DOCUMENT_PROPERTIES
    doc_props = struct.pack("<H", 1)
    doc_props += struct.pack("<HHHHHH", 1, 1, 1, 1, 1, 1)
    doc_props += struct.pack("<III", 0, 0, 0)
    records += make_record(HWPTAG_BEGIN + 0, 0, doc_props)
    # HWPTAG_ID_MAPPINGS
    id_map = struct.pack("<18i", 0,1,0,0,1,1,0,0,0,1,1,0,0,0,0,0,0,0)
    records += make_record(HWPTAG_BEGIN + 1, 0, id_map)
    # HWPTAG_FACE_NAME
    name = "나눔고딕"
    enc = name.encode("utf-16-le")
    face = struct.pack("<B", 0) + struct.pack("<H", len(name)) + enc
    face += struct.pack("<B", 0) + struct.pack("<H", 0) + b"\x00" * 10 + struct.pack("<H", 0)
    records += make_record(HWPTAG_BEGIN + 3, 1, face)
    # HWPTAG_CHAR_SHAPE (72 bytes)
    cs = struct.pack("<7H", *([0]*7))
    cs += struct.pack("<7B", *([100]*7))
    cs += struct.pack("<7b", *([0]*7))
    cs += struct.pack("<7B", *([100]*7))
    cs += struct.pack("<7b", *([0]*7))
    cs += struct.pack("<i", 1000)
    cs += struct.pack("<I", 0)
    cs += struct.pack("<bb", 0, 0)       # 그림자 간격 x2 = 2
    cs += struct.pack("<I", 0)           # 글자색         = 4
    cs += struct.pack("<I", 0)           # 밑줄색         = 4
    cs += struct.pack("<I", 0x00808080)  # 음영색         = 4
    cs += struct.pack("<I", 0)           # 그림자색       = 4
    cs += struct.pack("<H", 0)           # 글자테두리ID   = 2
    cs += struct.pack("<H", 0)           # padding        = 2
    # 14+7+7+7+7+4+4+2+4+4+4+4+2+2 = 72
    assert len(cs) == 72, f"cs={len(cs)}"
    records += make_record(HWPTAG_BEGIN + 5, 1, cs)
    # HWPTAG_PARA_SHAPE (54 bytes)
    ps  = struct.pack("<I",    0)          # 속성1          =  4
    ps += struct.pack("<iiii", 0,0,0,0)   # 여백/들여쓰기  = 16
    ps += struct.pack("<ii",   0,0)        # 문단간격       =  8
    ps += struct.pack("<i",    200)        # 줄간격(구버전) =  4
    ps += struct.pack("<HHH",  0,0,0)     # 탭/번호/테두리 =  6
    ps += struct.pack("<hhhh", 0,0,0,0)   # 테두리간격     =  8
    ps += struct.pack("<II",   0,200)      # 속성2, 줄간격  =  8
    # 합계: 4+16+8+4+6+8+8 = 54
    assert len(ps) == 54, f"ps={len(ps)}"
    records += make_record(HWPTAG_BEGIN + 9, 1, ps)
    # HWPTAG_STYLE
    sname = "바탕글"
    senc = sname.encode("utf-16-le")
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
        return StreamingResponse(
            io.BytesIO(hwp),
            media_type="application/x-hwp",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}.hwp"',
                "X-Pages": str(total),
                "Access-Control-Expose-Headers": "X-Pages,Content-Disposition",
            }
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"변환 오류:\n{traceback.format_exc()}")
        raise HTTPException(500, f"변환 오류: {str(e)}")
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
        return StreamingResponse(
            io.BytesIO(hwp),
            media_type="application/x-hwp",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}.hwp"',
                "X-Pages": str(total),
                "Access-Control-Expose-Headers": "X-Pages,Content-Disposition",
            }
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"변환 오류:\n{traceback.format_exc()}")
        raise HTTPException(500, f"변환 오류: {str(e)}")

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
        return StreamingResponse(
            io.BytesIO(hwp),
            media_type="application/x-hwp",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}.hwp"',
                "X-Pages": str(total),
                "Access-Control-Expose-Headers": "X-Pages,Content-Disposition",
            }
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"변환 오류:\n{traceback.format_exc()}")
        raise HTTPException(500, f"변환 오류: {str(e)}")
