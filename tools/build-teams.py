#!/usr/bin/env python3
"""Build a deterministic Teams tab package from the Outlook logo, monochrome routing geometry and manifest (requires Pillow)."""
from pathlib import Path
import struct,zlib,zipfile
ROOT=Path(__file__).resolve().parents[1]

def icons(folder):
    # Reuse the Outlook artwork; only adapt its size for the Teams package.
    from PIL import Image, ImageDraw
    with Image.open(ROOT/'assets/icon-300.png') as image:
        image.convert('RGBA').resize((192,192),Image.Resampling.LANCZOS).save(folder/'color.png')
    # Small monochrome counterpart: legislative document routed to three people.
    image=Image.new('RGBA',(256,256),(255,255,255,0));draw=ImageDraw.Draw(image)
    white=(255,255,255,255);clear=(255,255,255,0)
    draw.rounded_rectangle((22,92,133,225),radius=12,fill=white)
    draw.rectangle((38,145,115,154),fill=clear);draw.rectangle((38,170,115,179),fill=clear)
    draw.polygon([(36,65),(78,30),(120,65)],fill=white)
    draw.rectangle((32,69,124,78),fill=white)
    for x in [43,72,101]: draw.rectangle((x,83,x+12,112),fill=white)
    draw.line([(130,151),(154,151),(154,62),(179,62)],fill=white,width=9)
    draw.line([(154,151),(179,151)],fill=white,width=9)
    draw.line([(154,151),(154,222),(179,222)],fill=white,width=9)
    for y in [62,142,222]:
        draw.ellipse((177,y-23,223,y+23),outline=white,width=7)
        draw.ellipse((194,y-13,206,y-1),fill=white)
        draw.ellipse((189,y+3,211,y+17),fill=white)
    image.resize((32,32),Image.Resampling.LANCZOS).save(folder/'outline.png')

def build():
    folder=ROOT/'teams';folder.mkdir(exist_ok=True)
    icons(folder)
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
