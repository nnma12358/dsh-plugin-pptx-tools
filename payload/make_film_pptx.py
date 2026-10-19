#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_film_pptx.py — 把"图片/"下的所有照片做成"电影胶卷"风格幻灯片.

设计:
  * 16:9 黑底幻灯片, 上下两条黑色齿孔带 (白色圆角矩形 = 电影胶卷齿孔)
  * 每张照片按比例居中放在两带之间, 像一帧 35mm 胶片
  * 幻灯片自动放映: 轮盘(wheel)转场 + 定时 4s 自动切换 + 循环轮播

用法:
  ./venv/bin/python make_film_pptx.py [输出.pptx]
"""
import json, os, subprocess, sys, glob

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC_DIR = os.environ.get('PPTX_SRC_DIR') or os.path.join(ROOT, '图片')
ASSET_DIR = os.path.join(HERE, 'out', 'film-assets')
OUT_DEFAULT = os.path.join(ROOT, '图片集-电影胶卷.pptx')

SLIDE_W = 13.333
SLIDE_H = 7.5
BAND_H = 0.6                 # 上下齿孔带高度
HOLE_W, HOLE_H, HOLE_GAP = 0.30, 0.17, 0.34   # 齿孔尺寸/间距
MAX_DIM = 2000               # 预处理后最长边像素
JPEG_Q = 88
ADVANCE_SEC = 4.0            # 自动轮播间隔


def preprocess():
    """复制/压缩图片到 ASSET_DIR, 按文件名排序, 返回路径列表."""
    os.makedirs(ASSET_DIR, exist_ok=True)
    files = sorted(glob.glob(os.path.join(SRC_DIR, '*.jpg')) +
                   glob.glob(os.path.join(SRC_DIR, '*.jpeg')) +
                   glob.glob(os.path.join(SRC_DIR, '*.png')))
    assert files, '图片目录为空: %s' % SRC_DIR
    out = []
    for i, f in enumerate(files):
        from PIL import Image
        im = Image.open(f)
        im = ImageOps.exif_transpose(im).convert('RGB')
        if max(im.size) > MAX_DIM:
            im.thumbnail((MAX_DIM, MAX_DIM), Image.LANCZOS)
        dst = os.path.join(ASSET_DIR, '%02d.jpg' % (i + 1))
        im.save(dst, 'JPEG', quality=JPEG_Q, optimize=True)
        out.append(dst)
    return out


def hole_row(y_center):
    """一行齿孔: 返回围绕 y_center 居中的齿孔 rect 规格列表."""
    n = int((SLIDE_W - 0.4 + HOLE_GAP) // (HOLE_W + HOLE_GAP))
    total = n * HOLE_W + (n - 1) * HOLE_GAP
    x0 = (SLIDE_W - total) / 2.0
    specs = []
    for i in range(n):
        specs.append({'kind': 'rect', 'rounded': True, 'radius': 0.4,
                      'x': round(x0 + i * (HOLE_W + HOLE_GAP), 4),
                      'y': round(y_center - HOLE_H / 2.0, 4),
                      'w': HOLE_W, 'h': HOLE_H,
                      'fill': {'color': 'F5F5F5'}})
    return specs


def slide_spec(img_path, idx, total):
    from PIL import Image
    im = Image.open(img_path)
    w, h = im.size
    area_x, area_y = 0.5, BAND_H + 0.15
    area_w, area_h = SLIDE_W - 1.0, SLIDE_H - 2 * BAND_H - 0.30
    scale = min(area_w / w, area_h / h)
    pw, ph = w * scale, h * scale
    x = (SLIDE_W - pw) / 2.0
    y = (SLIDE_H - ph) / 2.0
    shapes = []
    # 上下黑色带
    shapes.append({'kind': 'rect', 'rounded': False, 'x': 0, 'y': 0,
                   'w': SLIDE_W, 'h': BAND_H, 'fill': {'color': '000000'}})
    shapes.append({'kind': 'rect', 'rounded': False, 'x': 0, 'y': SLIDE_H - BAND_H,
                   'w': SLIDE_W, 'h': BAND_H, 'fill': {'color': '000000'}})
    # 齿孔
    shapes += hole_row(BAND_H / 2.0)
    shapes += hole_row(SLIDE_H - BAND_H / 2.0)
    # 照片
    shapes.append({'kind': 'image', 'path': img_path,
                   'x': round(x, 4), 'y': round(y, 4),
                   'w': round(pw, 4), 'h': round(ph, 4)})
    return {
        'bg': {'type': 'solid', 'color': '000000'},
        'shapes': shapes,
        'transition': {'type': 'wheel', 'spd': 'med',
                       'advanceAfterSec': ADVANCE_SEC, 'advanceOnClick': True},
        'notes': '%02d / %d  %s' % (idx, total, os.path.basename(img_path)),
    }


def build(out_path, assets):
    spec = {
        'slideWidth': SLIDE_W,
        'slideHeight': SLIDE_H,
        'show': {'loop': True, 'show': 'all', 'useTimings': True},
        'slides': [],
    }
    total = len(assets)
    for i, img in enumerate(assets, start=1):
        spec['slides'].append(slide_spec(img, i, total))
    spec_path = os.path.join(HERE, 'out', 'film-spec.json')
    os.makedirs(os.path.dirname(spec_path), exist_ok=True)
    json.dump(spec, open(spec_path, 'w'), ensure_ascii=False, indent=1)

    from pptx_cli import build as build_pkg
    res = build_pkg(spec, out_path)
    res['spec'] = spec_path
    return res


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else OUT_DEFAULT
    assets = preprocess()
    res = build(out, assets)
    print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    # PIL ImageOps 在 preprocess 内使用时再导入
    import PIL.ImageOps as _io  # noqa
    from PIL import Image, ImageOps
    main()
