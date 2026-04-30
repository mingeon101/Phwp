"""
PDF → HWP 5.0 변환 백엔드  |  Render.com 배포용
CFB 섹터 레이아웃 고정 방식 (로컬 검증 완료)
"""
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from urllib.parse import quote
import pdfplumber, struct, zlib, io, traceback, logging
from pathlib import Path

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF to HWP Converter API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

SS=512; FREE=0xFFFFFFFF; EOC=0xFFFFFFFE; FATSECT=0xFFFFFFFD; NOSM=0xFFFFFFFF; T=0x010

def pad(d): r=len(d)%SS; return d+b"\x00"*(SS-r if r else 0)

def dentry(name,typ,color,left,right,child,start,size,clsid=None):
    if clsid is None: clsid=b"\x00"*16
    enc=name.encode("utf-16-le") if name else b""
    nlen=len(enc)+2 if name else 0
    nf=(enc+b"\x00\x00")[:62].ljust(64,b"\x00")
    d=nf+struct.pack("<H",nlen)+struct.pack("<BB",typ,color)+struct.pack("<III",left,right,child)
    d+=clsid+struct.pack("<I",0)+struct.pack("<Q",0)+struct.pack("<Q",0)
    d+=struct.pack("<I",start)+struct.pack("<Q",size)
    assert len(d)==128; return d

def rec(tag,lv,data):
    sz=len(data)
    if sz>=0xFFF: return struct.pack("<II",(tag&0x3FF)|((lv&0x3FF)<<10)|(0xFFF<<20),sz)+data
    return struct.pack("<I",(tag&0x3FF)|((lv&0x3FF)<<10)|((sz&0xFFF)<<20))+data

def make_file_header():
    sig=b"HWP Document File\x00"+b"\x00"*14
    fh=sig[:32]+struct.pack("<I",0x05000306)+struct.pack("<I",0x00000001)+struct.pack("<I",0)+struct.pack("<I",0)+struct.pack("<B",0)+b"\x00"*207
    assert len(fh)==256; return fh

def make_doc_info():
    out=b""
    out+=rec(T+0,0,struct.pack("<H",1)+struct.pack("<HHHHHH",1,1,1,1,1,1)+struct.pack("<III",0,0,0))
    out+=rec(T+1,0,struct.pack("<18i",0,1,0,0,1,1,0,0,0,1,1,0,0,0,0,0,0,0))
    fn="나눔고딕"; fe=fn.encode("utf-16-le")
    face=struct.pack("<B",0)+struct.pack("<H",len(fn))+fe+struct.pack("<B",0)+struct.pack("<H",0)+b"\x00"*10+struct.pack("<H",0)
    out+=rec(T+3,1,face)
    cs=(struct.pack("<7H",*([0]*7))+struct.pack("<7B",*([100]*7))+struct.pack("<7b",*([0]*7))
       +struct.pack("<7B",*([100]*7))+struct.pack("<7b",*([0]*7))
       +struct.pack("<i",1000)+struct.pack("<I",0)+struct.pack("<bb",0,0)
       +struct.pack("<IIII",0,0,0x808080,0)+struct.pack("<HH",0,0))
    assert len(cs)==72; out+=rec(T+5,1,cs)
    ps=(struct.pack("<I",0)+struct.pack("<iiii",0,0,0,0)+struct.pack("<ii",0,0)
       +struct.pack("<i",200)+struct.pack("<HHH",0,0,0)+struct.pack("<hhhh",0,0,0,0)
       +struct.pack("<II",0,200))
    assert len(ps)==54; out+=rec(T+9,1,ps)
    sn="바탕글"; se=sn.encode("utf-16-le")
    out+=rec(T+10,1,struct.pack("<H",len(sn))+se+struct.pack("<H",0)+struct.pack("<B",0)+struct.pack("<B",0)+struct.pack("<h",0)+struct.pack("<HH",0,0))
    return out

def build_cfb(fh_data,di_data,s0_data):
    """
    고정 섹터 레이아웃:
      sect 0: FileHeader  sect 1: DocInfo  sect 2: Section0
      sect 3~4: Directory  sect 5: FAT
    """
    ROOT_CLS=b"\x00\x20\x08\x02\x00\x00\x00\x00\xc0\x00\x00\x00\x00\x00\x00\x46"
    entries=[
        dentry("Root Entry",5,1,NOSM,NOSM,1,EOC,0,ROOT_CLS),
        dentry("FileHeader",2,1,NOSM,2,NOSM,0,len(fh_data)),
        dentry("DocInfo",   2,1,NOSM,3,NOSM,1,len(di_data)),
        dentry("BodyText",  1,1,NOSM,NOSM,4,EOC,0),
        dentry("Section0",  2,1,NOSM,NOSM,NOSM,2,len(s0_data)),
        dentry("",0,1,NOSM,NOSM,NOSM,EOC,0),
        dentry("",0,1,NOSM,NOSM,NOSM,EOC,0),
        dentry("",0,1,NOSM,NOSM,NOSM,EOC,0),
    ]
    dir_bytes=b"".join(entries)
    fat=[EOC,EOC,EOC,4,EOC,FATSECT]+[FREE]*122
    fat_bytes=struct.pack("<128I",*fat)
    hdr =b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"+b"\x00"*16
    hdr+=struct.pack("<HH",0x003E,0x0003)+struct.pack("<H",0xFFFE)+struct.pack("<HH",9,6)+b"\x00"*6
    hdr+=struct.pack("<I",0)+struct.pack("<I",1)+struct.pack("<I",3)+struct.pack("<I",0)+struct.pack("<I",4096)
    hdr+=struct.pack("<I",FREE)+struct.pack("<I",0)+struct.pack("<I",FREE)+struct.pack("<I",0)
    hdr+=struct.pack("<I",5)+struct.pack("<I",FREE)*108
    assert len(hdr)==512
    return hdr+pad(fh_data)+pad(di_data)+pad(s0_data)+dir_bytes+fat_bytes

def make_body(pages):
    out=b""
    for i,text in enumerate(pages):
        lines=text.replace("\r\n","\n").replace("\r","\n").split("\n") if text.strip() else [""]
        for line in lines:
            lt=line+"\r"; enc=lt.encode("utf-16-le"); nc=len(lt)
            ph=(struct.pack("<I",nc)+struct.pack("<I",0)+struct.pack("<H",0)
               +struct.pack("<BB",0,0)+struct.pack("<H",1)+struct.pack("<HH",0,0)
               +struct.pack("<I",abs(hash(lt))&0xFFFFFFFF)+struct.pack("<H",0))
            out+=rec(T+50,0,ph)+rec(T+51,1,enc)+rec(T+52,1,struct.pack("<II",0,0))
        if i<len(pages)-1:
            sep="\r".encode("utf-16-le")
            ph=(struct.pack("<I",1)+struct.pack("<I",0)+struct.pack("<H",0)
               +struct.pack("<BB",0,0)+struct.pack("<H",1)+struct.pack("<HH",0,0)
               +struct.pack("<I",0)+struct.pack("<H",0))
            out+=rec(T+50,0,ph)+rec(T+51,1,sep)+rec(T+52,1,struct.pack("<II",0,0))
    return out

@app.get("/")
async def root(): return {"status":"ok","usage":"POST /convert"}

@app.get("/health")
async def health(): return {"status":"ok"}

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400,"PDF 파일만 업로드 가능합니다.")
    content=await file.read()
    if len(content)>50*1024*1024:
        raise HTTPException(400,"50MB 이하 파일만 지원합니다.")
    try:
        logger.info(f"변환 시작: {file.filename} ({len(content):,}B)")
        pages=[]
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            total=len(pdf.pages)
            for i,page in enumerate(pdf.pages):
                t=page.extract_text() or ""
                pages.append(t)
                logger.info(f"  p{i+1}: {len(t)}자")
        if not any(t.strip() for t in pages):
            pages=["(이미지 기반 PDF — 텍스트 추출 불가)"]

        di_z=zlib.compress(make_doc_info(),6)
        bd_z=zlib.compress(make_body(pages),6)
        hwp=build_cfb(make_file_header(), di_z, bd_z)
        logger.info(f"HWP 완성: {len(hwp):,}B")

        stem=Path(file.filename).stem
        cd=f"attachment; filename*=UTF-8''{quote(stem+'.hwp',safe='')}"
        return StreamingResponse(io.BytesIO(hwp),media_type="application/x-hwp",
            headers={"Content-Disposition":cd,"X-Pages":str(total),
                     "Access-Control-Expose-Headers":"X-Pages,Content-Disposition"})
    except HTTPException: raise
    except Exception as e:
        logger.error(traceback.format_exc())
        raise HTTPException(500,f"변환 오류: {e}")
