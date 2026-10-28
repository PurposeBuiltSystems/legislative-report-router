#!/usr/bin/env python3
"""Build a deterministic Teams tab package from repo-native icon geometry and manifest."""
from pathlib import Path
import struct,zlib,zipfile
ROOT=Path(__file__).resolve().parents[1]

def png(size, background):
    pixels=[list(background) for _ in range(size*size)]
    def draw(x,y):
        x=x/size*192;y=y/size*192
        roof=28<=y<=68 and abs(x-96)<=(y-28)*1.5
        columns=82<=y<=139 and any(a<=x<=a+14 for a in [48,89,130])
        bars=(36<=x<=156 and (68<=y<=78 or 146<=y<=160))
        return roof or columns or bars
    for y in range(size):
        for x in range(size):
            if draw(x,y):pixels[y*size+x]=[255,255,255,255]
    raw=b''.join(b'\x00'+bytes(sum(pixels[y*size:(y+1)*size],[])) for y in range(size))
    def chunk(kind,data):return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data)&0xffffffff)
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',size,size,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(raw,9))+chunk(b'IEND',b'')

def build():
    folder=ROOT/'teams';folder.mkdir(exist_ok=True)
    (folder/'color.png').write_bytes(png(192,[15,108,189,255]))
    (folder/'outline.png').write_bytes(png(32,[0,0,0,0]))
    package=folder/'Legislative-Work-Queue.zip'
    with zipfile.ZipFile(package,'w',compression=zipfile.ZIP_DEFLATED) as archive:
        for name in ['manifest.json','color.png','outline.png']:
            info=zipfile.ZipInfo(name,date_time=(2026,10,28,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED;info.external_attr=0o100644<<16;archive.writestr(info,(folder/name).read_bytes())
    print(package)
    handoff=folder/'Legislative-Router-Admin-Handoff.zip'
    with zipfile.ZipFile(handoff,'w',compression=zipfile.ZIP_DEFLATED) as archive:
        entries={'README.txt':folder/'admin-request.txt','permission-request.json':folder/'permission-request.json','Legislative-Work-Queue.zip':package,'Contact-sync-requirements.md':ROOT/'docs/contact-sync-admin.md'}
        for name,source in entries.items():
            info=zipfile.ZipInfo(name,date_time=(2026,10,28,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED;info.external_attr=0o100644<<16;archive.writestr(info,source.read_bytes())
    print(handoff)
if __name__=='__main__':build()
