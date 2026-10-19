#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_film_carousel.py — 单页"电影胶卷"向右滑动轮播幻灯片.

设计:
  * 所有照片放在同一页 (1 张幻灯片, 16:9 黑底)
  * 每张照片缩到 35mm 胶卷画幅大小 (约 1.02" x 0.77", 4:3), 全部在一卷胶带上:
    中间照片, 上下两条齿孔带 (白色圆角矩形)
  * 幻灯片放映时整卷胶带以"向右滑动"方式轮播:
    - 胶带内容以 11 张照片为一周期 (重复 2 次 = 24 帧)
    - 每 7 秒向右平移一个完整周期 (11 x 1.2"), 然后无缝回绕
    - 动画为 PowerPoint 原生 timing XML: afterEffect 自动开始 + repeatCount 无限循环

用法:
  ./venv/bin/python make_film_carousel.py [输出.pptx]
"""
import os, sys, glob, json
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC_DIR = os.environ.get('PPTX_SRC_DIR') or os.path.join(ROOT, '图片')
ASSET_DIR = os.path.join(HERE, 'out', 'film-assets')
OUT_DEFAULT = os.path.join(ROOT, '图片集-电影胶卷.pptx')

SW = 13.333          # 幻灯片宽 (英寸)
SH = 7.5             # 幻灯片高
FRAME_W = 1.02       # 胶卷画幅宽
FRAME_H = 0.765      # 胶卷画幅高 (4:3)
PITCH = 1.20         # 帧距 (画幅 + 间隔)
HOLE_W, HOLE_H = 0.14, 0.095      # 齿孔
HOLE_OFF = 0.27                   # 每帧 2 个齿孔相对画幅中心的偏移
RAIL = 0.22                       # 齿孔行距胶带边缘
MAX_DIM = 2000
JPEG_Q = 88
CYCLE_SEC = 7.0                   # 每周期秒数 (向右滑动一个完整周期)


def preprocess():
    os.makedirs(ASSET_DIR, exist_ok=True)
    files = sorted(glob.glob(os.path.join(SRC_DIR, '*.jpg')) +
                   glob.glob(os.path.join(SRC_DIR, '*.jpeg')) +
                   glob.glob(os.path.join(SRC_DIR, '*.png')))
    assert files, '图片目录为空: %s' % SRC_DIR
    out = []
    for i, f in enumerate(files):
        im = Image.open(f)
        im = ImageOps.exif_transpose(im).convert('RGB')
        if max(im.size) > MAX_DIM:
            im.thumbnail((MAX_DIM, MAX_DIM), Image.LANCZOS)
        dst = os.path.join(ASSET_DIR, '%02d.jpg' % (i + 1))
        im.save(dst, 'JPEG', quality=JPEG_Q, optimize=True)
        out.append(dst)
    return out


def build(assets, out_path):
    n = len(assets)
    strip_frames = 2 * n + 2          # 照片周期重复两次 + 2 帧余量
    body_w = strip_frames * PITCH
    # 向右滑动: 胶带左缘必须一直 <= 0, 且滑动一整周期后回绕无缝
    body_x0 = -(n * PITCH) - 0.30
    body_top = (SH - (FRAME_H + 2 * RAIL + 2 * HOLE_H)) / 2.0
    body_h = FRAME_H + 2 * (RAIL + HOLE_H)
    body_bottom = body_top + body_h
    cy = SH / 2.0                     # 画幅垂直中心

    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 黑色背景
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

    grp = slide.shapes.add_group_shape()
    gs = grp.shapes

    def add_rect(x, y, w, h, color):
        sp = gs.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sp.fill.solid()
        sp.fill.fore_color.rgb = RGBColor.from_string(color)
        sp.line.fill.background()
        sp.shadow.inherit = False
        return sp

    # 胶带本体 (黑)
    add_rect(body_x0, body_top, body_w, body_h, '000000')

    # 齿孔 + 照片
    for i in range(strip_frames):
        fx = body_x0 + i * PITCH + (PITCH - FRAME_W) / 2.0
        cx = fx + FRAME_W / 2.0
        for dx in (-HOLE_OFF, HOLE_OFF):
            add_rect(cx + dx - HOLE_W / 2.0, body_top + RAIL, HOLE_W, HOLE_H, 'F5F5F5')
            add_rect(cx + dx - HOLE_W / 2.0, body_bottom - RAIL - HOLE_H, HOLE_W, HOLE_H, 'F5F5F5')
        gs.add_picture(assets[i % n], Inches(fx), Inches(cy - FRAME_H / 2.0),
                       Inches(FRAME_W), Inches(FRAME_H))

    # ---- 动画: 向右滑动一整个周期, 自动开始, 无限循环 ----
    spid = grp._element.find('.//' + qn('p:cNvPr')).get('id')
    dx = n * PITCH / SW               # 平移量 (幻灯片宽度的比例)
    dur = str(int(CYCLE_SEC * 1000))
    tmp = """
<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
<p:childTnLst><p:seq concurrent="1" nextAc="seek">
<p:cTn id="2" dur="indefinite" nodeType="mainSeq">
<p:childTnLst>
<p:par><p:cTn id="3" fill="hold">
<p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:par><p:cTn id="4" fill="hold">
<p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst><p:par>
<p:cTn id="5" presetID="2" presetClass="entr" presetSubtype="3" fill="hold" grpId="0"
       nodeType="afterEffect" repeatCount="indefinite">
<p:stCondLst><p:cond delay="0"/></p:stCondLst>
<p:childTnLst>
<p:anim calcmode="lin" valueType="num">
<p:cBhvr additive="base"><p:cTn id="7" dur="%DUR" fill="hold"/>
<p:tgtEl><p:spTgt spid="%ID"/></p:tgtEl>
<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>
<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>
<p:tav tm="100000"><p:val><p:strVal val="#ppt_x+%DX"/></p:val></p:tav></p:tavLst></p:anim>
<p:anim calcmode="lin" valueType="num">
<p:cBhvr additive="base"><p:cTn id="8" dur="%DUR" fill="hold"/>
<p:tgtEl><p:spTgt spid="%ID"/></p:tgtEl>
<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>
<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav>
<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>
</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
</p:childTnLst></p:cTn>
<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>
""".replace('%DUR', dur).replace('%DX', '%.5f' % dx).replace('%ID', spid)
    import lxml.etree as ET
    timing = ET.fromstring(tmp)
    slide._element.append(timing)

    # 演示设置: 自动放映 + 循环
    presEl = prs._element
    showPr = presEl.makeelement(qn('p:showPr'), {})
    showPr.set('loop', '1')
    showPr.set('showAnimation', '1')
    showPr.set('useTimings', '1')
    sldAll = presEl.makeelement(qn('p:sldAll'), {})
    showPr.append(sldAll)
    sldSz = presEl.find(qn('p:sldSz'))
    if sldSz is not None:
        sldSz.addnext(showPr)
    else:
        presEl.append(showPr)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or '.', exist_ok=True)
    prs.save(out_path)
    return {'out': out_path, 'photos': n, 'frames': strip_frames,
            'periodIn': n * PITCH, 'cycleSec': CYCLE_SEC}


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else OUT_DEFAULT
    assets = preprocess()
    res = build(assets, out)
    print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
