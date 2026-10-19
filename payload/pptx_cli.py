#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptx_cli.py — read and build PowerPoint (.pptx) files.

Subcommands:
  inspect <file.pptx>                 -> JSON summary (slides, shapes, text+font info,
                                         backgrounds, transitions, animations, notes,
                                         theme, show settings, media)
  build <spec.json> <out.pptx> [--to-json]  -> build a pptx from a structured spec;
                                         prints JSON {out, slideCount, warnings}
  demo <out.pptx>                     -> generate a demo file for testing
"""
import sys, os, json, zipfile, re, copy, argparse, shutil, tempfile
import xml.etree.ElementTree as ET
from io import BytesIO

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

def q(tag):  # 'p:showPr' -> '{ns}showPr'
    pfx, _, local = tag.partition(':')
    ns = {'p': P_NS, 'a': A_NS, 'r': R_NS, 'pr': RELS_NS}.get(pfx)
    return '{%s}%s' % (ns, local) if ns else tag

def local(tag):
    return tag.split('}')[-1] if '}' in tag else tag

# ---------------------------------------------------------------------------
# helpers for color / font readout
# ---------------------------------------------------------------------------
def color_of(font_color):
    """Return readable color info {'rgb': 'FF8800'} or {'theme': 'accent1'}."""
    try:
        try:
            rgb = font_color.rgb
            return {'rgb': str(rgb)}
        except Exception:
            pass
        try:
            tc = font_color.theme_color
            return {'theme': str(tc)}
        except Exception:
            return {}
    except Exception:
        return {}

def color_of_fill(fill):
    """Best-effort fill description."""
    try:
        ft = str(fill.type)
        if 'SOLID' in ft.upper() or ft == 'solid':
            info = color_of(fill.fore_color)
            return {'type': 'solid', **info}
        if 'BACKGROUND' in ft.upper() or ft == 'background':
            return {'type': 'background'}
        if 'PICTURE' in ft.upper() or ft == 'picture':
            return {'type': 'picture'}
        if 'GRADIENT' in ft.upper() or ft == 'gradient':
            try:
                stops = []
                for st in fill.gradient_stops:
                    stops.append({'pos': float(st.position), **color_of(st.color)})
                return {'type': 'gradient', 'angle': getattr(fill, 'gradient_angle', None), 'stops': stops}
            except Exception:
                return {'type': 'gradient'}
        return {'type': ft}
    except Exception:
        return {'type': 'unknown'}

# ---------------------------------------------------------------------------
# INSPECT
# ---------------------------------------------------------------------------
def inspect(path):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.oxml.ns import qn as pqn

    prs = Presentation(path)
    out = {}
    out['file'] = os.path.basename(path)
    try:
        out['fileSize'] = os.path.getsize(path)
    except Exception:
        pass

    # ---- presentation level
    pres = {}
    pres['slideWidthIn'] = round(prs.slide_width / 914400, 3)
    pres['slideHeightIn'] = round(prs.slide_height / 914400, 3)
    pres['slideCount'] = len(prs.slides._sldIdLst)

    # show settings from raw XML
    presEl = prs._element
    showPr = presEl.find(q('p:showPr'))
    if showPr is not None:
        attr = {k.split('}')[-1]: v for k, v in showPr.attrib.items()}
        attr = {k: (v == '1') if v in ('0', '1') else v for k, v in attr.items()}
        show = {'attrs': attr}
        for child in showPr:
            tag = local(child.tag)
            cattr = {k.split('}')[-1]: v for k, v in child.attrib.items()}
            show.setdefault('children', []).append({'kind': tag, **cattr})
        pres['show'] = show
    pres['slideShowDefaults'] = {
        'advanceOnClick': presEl.get('advClick'),
        'autoAdvance': presEl.get('autoAdvance'),
        'advanceTimeMs': presEl.get('advanceTm'),
    }

    # theme (fonts + colors)
    try:
        from pptx.oxml.ns import nsmap
        theme = read_theme(prs)
        pres['theme'] = theme
    except Exception as e:
        pres['theme'] = {'error': str(e)}

    # masters / layouts
    try:
        masters = []
        for m in prs.slide_masters:
            masters.append(getattr(m, 'name', None) or 'slideMaster%d' % (m.part.partname.split('slideMaster')[-1].split('.')[0]))
        pres['masters'] = masters
    except Exception:
        pass
    try:
        layouts = []
        for s in prs.slides:
            lay = s.slide_layout
            layouts.append(lay.name if lay is not None else '?')
        pres['layoutPerSlide'] = layouts
    except Exception:
        pass

    # media
    try:
        media = []
        z = zipfile.ZipFile(path)
        for name in z.namelist():
            if name.startswith('ppt/media/'):
                media.append({'file': os.path.basename(name), 'size': z.getinfo(name).file_size})
        pres['media'] = media
    except Exception:
        pass

    out['presentation'] = pres

    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        info = slide_info(slide, idx, path)
        slides.append(info)
    out['slides'] = slides
    out['summary'] = {
        'slides': len(slides),
        'textShapes': sum(1 for s in slides for sh in s['shapes'] if sh['type'] in ('text', 'placeholder', 'table')),
        'pictures': sum(1 for s in slides for sh in s['shapes'] if sh['type'] == 'picture'),
        'shapesWithAnimation': sum(1 for s in slides for a in s['animations'] if a.get('targetShape')),
        'slidesWithTransition': [s['index'] for s in slides if s.get('transition')],
    }
    return out

def read_theme(prs):
    """Read theme fonts/colors from the package zip."""
    from pptx.oxml.ns import nsmap
    themepart = None
    for part in prs.part.package.iter_parts():
        name = str(part.partname)
        if name.startswith('/ppt/theme/') and name.endswith('.xml'):
            themepart = part
            break
    if themepart is None:
        return {}
    try:
        root = ET.fromstring(themepart.blob)
    except Exception:
        return {}
    res = {}
    t = {}
    for child in root:
        tag = local(child.tag)
        if tag == 'themeElements':
            for te in child:
                tte = local(te.tag)
                if tte == 'clrScheme':
                    colors = {}
                    for c in te:
                        cname = local(c.tag)
                        for v in c:
                            vtag = local(v.tag)
                            if vtag == 'srgbClr':
                                colors[cname] = v.get('val')
                            elif vtag == 'sysClr':
                                colors[cname] = v.get('lastClr')
                    res['colors'] = colors
                elif tte == 'fontScheme':
                    fonts = {}
                    for fs in te:
                        fst = local(fs.tag)
                        if fst in ('majorFont', 'minorFont'):
                            lat = fs.find(q('a:latin'))
                            ea = fs.find(q('a:ea'))
                            fonts[fst] = {
                                'latin': lat.get('typeface') if lat is not None else None,
                                'ea': ea.get('typeface') if ea is not None else None,
                            }
                    res['fonts'] = fonts
    return res

def slide_info(slide, idx, path):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.oxml.ns import qn as pqn
    info = {'index': idx}
    try:
        lay = slide.slide_layout
        info['layout'] = lay.name if lay is not None else None
    except Exception:
        pass

    # background
    try:
        bg = slide.background
        fill_desc = color_of_fill(bg.fill) if bg else None
    except Exception:
        fill_desc = None
    # raw XML read: more reliable for gradient/picture
    try:
        cSld = slide._element.find(q('p:cSld'))
        bgEl = cSld.find(q('p:bg')) if cSld is not None else None
        if bgEl is not None:
            raw = parse_raw_bg(bgEl, slide)
            if raw is not None:
                fill_desc = raw
    except Exception:
        pass
    info['background'] = fill_desc

    # shapes
    shapes = []
    spids = {}
    spTree = slide._element.find('.//' + q('p:spTree'))
    if spTree is not None:
        for s in spTree:
            tag = local(s.tag)
            if tag in ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'):
                cNvPr = s.find('.//' + q('p:cNvPr')) if tag in ('sp', 'pic', 'graphicFrame', 'cxnSp') else s.find('.//' + q('a:cNvPr'))
                sid = cNvPr.get('id') if cNvPr is not None else s.get('id')
                spids[sid] = (cNvPr.get('name') if cNvPr is not None else None, tag)

    for sh in slide.shapes:
        shapes.append(shape_info(sh, slide, path))
    info['shapes'] = shapes

    # transition (raw xml)
    transEl = slide._element.find(q('p:transition'))
    if transEl is not None:
        tr = {k.split('}')[-1]: v for k, v in transEl.attrib.items()}
        for c in transEl:
            tr['effect'] = local(c.tag)
            tr.update({k.split('}')[-1]: v for k, v in c.attrib.items()})
        info['transition'] = tr

    # animations (raw xml timing tree)
    info['animations'] = parse_animations(slide, spids)

    # notes
    try:
        if slide.has_notes_slide:
            tf = slide.notes_slide.notes_text_frame
            if tf is not None and tf.text.strip():
                info['notes'] = tf.text
    except Exception:
        pass
    return info

def parse_raw_bg(bgEl, slide):
    """Parse p:bg XML into a readable dict (solid / gradient / picture)."""
    def visit(node):
        tag = local(node.tag)
        if tag == 'solidFill':
            return {'type': 'solid', **read_fill_color(node)}
        if tag == 'gradFill':
            stops = []
            for gs in node.findall(q('a:gsLst') + '/' + q('a:gs')):
                col = read_fill_color(gs)
                stops.append({'pos': int(gs.get('pos', 0)) / 100000, **col})
            lin = node.find(q('a:lin'))
            angle = int(lin.get('ang', 0)) / 60000 if lin is not None else None
            return {'type': 'gradient', 'angle': angle, 'stops': stops}
        if tag == 'blipFill':
            blip = node.find(q('a:blip'))
            if blip is not None:
                rid = blip.get(q('r:embed'))
                try:
                    target = slide.part.rels[rid].target_partname
                    return {'type': 'picture', 'image': str(target).split('/')[-1]}
                except Exception:
                    return {'type': 'picture'}
            return {'type': 'picture'}
        if tag in ('noFill',):
            return {'type': 'none'}
        # scheme-based fill via bgRef
        for c in node:
            r = visit(c)
            if r is not None:
                return r
        return None
    bgPr = bgEl.find(q('p:bgPr'))
    if bgPr is not None:
        for c in bgPr:
            r = visit(c)
            if r is not None:
                return r
    bgRef = bgEl.find(q('p:bgRef'))
    if bgRef is not None:
        scheme = bgRef.find(q('a:schemeClr'))
        return {'type': 'scheme', 'theme': scheme.get('val') if scheme is not None else bgRef.get('idx')}
    return None

def read_fill_color(el):
    """Read the color child of a fill node: srgbClr / schemeClr / sysClr."""
    for c in el:
        tag = local(c.tag)
        if tag == 'srgbClr':
            return {'rgb': c.get('val')}
        if tag == 'schemeClr':
            return {'theme': c.get('val')}
        if tag == 'sysClr':
            return {'rgb': c.get('lastClr'), 'sys': c.get('val')}
        # nested colorChoice / hsl / prstClr
        if tag == 'hslClr':
            return {'hsl': {k: v for k, v in c.attrib.items()}}
        if tag == 'prstClr':
            return {'preset': c.get('val')}
        inner = read_fill_color(c)
        if inner:
            return inner
    return {}

def shape_info(sh, slide, path):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.oxml.ns import qn as pqn
    t = str(sh.shape_type)
    d = {'type': t, 'name': sh.name}
    def dims(bbox):
        return {
            'xIn': round(bbox.left / 914400, 3) if bbox.left is not None else None,
            'yIn': round(bbox.top / 914400, 3) if bbox.top is not None else None,
            'wIn': round(bbox.width / 914400, 3) if bbox.width is not None else None,
            'hIn': round(bbox.height / 914400, 3) if bbox.height is not None else None,
        }
    try:
        d.update(dims(sh))
    except Exception:
        pass
    try:
        d['rotation'] = sh.rotation
    except Exception:
        pass

    try:
        if sh.has_text_frame:
            paras = []
            for p in sh.text_frame.paragraphs:
                runs = []
                for r in p.runs:
                    rd = {'text': r.text}
                    f = r.font
                    rd['font'] = f.name
                    if f.size:
                        rd['sizePt'] = f.size.pt
                    if f.bold:
                        rd['bold'] = True
                    if f.italic:
                        rd['italic'] = True
                    if f.underline:
                        rd['underline'] = True
                    col = color_of(f.color)
                    if col:
                        rd['color'] = col
                    # east-asian font
                    try:
                        rPr = r._r.find(qn('a:rPr'))
                        if rPr is not None:
                            ea = rPr.find(pqn('a:ea'))
                            if ea is not None:
                                rd['fontEa'] = ea.get('typeface')
                            lat = rPr.find(pqn('a:latin'))
                            if lat is not None:
                                rd['font'] = lat.get('typeface') or rd.get('font')
                    except Exception:
                        pass
                    try:
                        if r.hyperlink and r.hyperlink.address:
                            rd['link'] = r.hyperlink.address
                    except Exception:
                        pass
                    if rd['text']:
                        runs.append(rd)
                try:
                    align = p.alignment
                    palign = str(align) if align is not None else None
                except Exception:
                    palign = None
                paras.append({'text': p.text, 'align': palign, 'level': p.level, 'runs': runs})
            d['text'] = {'paragraphs': paras}
    except Exception:
        pass

    try:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                d['image'] = {'file': sh.image.filename, 'ext': sh.image.ext,
                              'size': (sh.image.size[0], sh.image.size[1])}
            except Exception:
                d['image'] = {'file': None}
            try:
                el = sh._element.find('.//' + q('a:blip'))
                if el is not None:
                    rid = el.get(q('r:embed'))
                    if rid:
                        tgt = slide.part.rels[rid].target_partname
                        d['image']['part'] = str(tgt).split('/')[-1]
            except Exception:
                pass
    except Exception:
        pass

    try:
        if sh.has_table:
            rows = []
            for row in sh.table.rows:
                rows.append([c.text for c in row.cells])
            d['table'] = {'rows': len(sh.table.rows), 'cols': len(sh.table.columns), 'data': rows}
    except Exception:
        pass

    try:
        if sh.has_chart:
            d['chart'] = {'type': str(sh.chart.chart_type)}
    except Exception:
        pass

    try:
        if sh.shape_type == MSO_SHAPE_TYPE.MEDIA:
            d['media'] = True
    except Exception:
        pass

    try:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            d['group'] = [shape_info(c, slide, path) for c in sh.shapes]
    except Exception:
        pass

    if not d.get('text'):
        try:
            if sh.has_text_frame and sh.text_frame.text.strip():
                d['text'] = {'paragraphs': [{'text': sh.text_frame.text}]}
        except Exception:
            pass
    return d

def parse_animations(slide, spids):
    """Extract a readable list of animation effects from p:timing."""
    timing = slide._element.find(q('p:timing'))
    if timing is None:
        return []
    anims = []
    click = 0
    for par in timing.iter(q('p:par')):
        cTn = None
        for child in par:
            if local(child.tag) == 'cTn':
                cTn = child
                break
        if cTn is None:
            continue
        ntype = cTn.get('nodeType', '')
        if ntype not in ('clickEffect', 'withEffect', 'afterEffect', 'effect'):
            continue
        if ntype == 'clickEffect':
            click += 1
        entry = {'click': click, 'nodeType': ntype}
        if cTn.get('presetID'):
            entry['presetID'] = cTn.get('presetID')
        if cTn.get('presetClass'):
            entry['presetClass'] = cTn.get('presetClass')
        if cTn.get('presetSubtype'):
            entry['presetSubtype'] = cTn.get('presetSubtype')
        if cTn.get('dur'):
            entry['durMs'] = cTn.get('dur')
        # delay
        for cond in cTn.iter(q('p:cond')):
            if cond.get('delay') and cond.get('delay') != 'indefinite':
                try:
                    entry['delayMs'] = int(cond.get('delay'))
                except Exception:
                    entry['delayMs'] = cond.get('delay')
            break
        # target shape
        spTgt = cTn.find('.//' + q('p:spTgt'))
        if spTgt is not None:
            spid = spTgt.get('spid')
            entry['targetShapeId'] = spid
            nm = spids.get(spid)
            if nm:
                entry['targetShape'] = nm[0]
                entry['targetKind'] = nm[1]
        # behavior type: first child of childTnLst
        ctl = cTn.find(q('p:childTnLst'))
        if ctl is not None and len(ctl):
            behavior = local(ctl[0].tag)
            entry['behavior'] = behavior
            # values
            if behavior == 'anim':
                tavs = [(int(t.get('tm', 0)), ''.join(t.findtext('.//' + q('p:strVal'), default='')) ) for t in ctl[0].findall(q('p:tavLst') + '/' + q('p:tav'))] if ctl[0].find(q('p:tavLst')) is not None else []
                if tavs:
                    entry['motion'] = tavs
        entry = {k: v for k, v in entry.items() if v not in (None, '', [])}
        anims.append(entry)
    return anims

# ---------------------------------------------------------------------------
# BUILD
# ---------------------------------------------------------------------------
def build(spec, out_path):
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.oxml.ns import qn

    warnings = []
    SLIDE_W = float(spec.get('slideWidth', 13.333))
    SLIDE_H = float(spec.get('slideHeight', 7.5))

    prs = Presentation()  # default template (4:3); we override size
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank = prs.slide_layouts[6]

    def hex_color(v, default='000000'):
        if v is None:
            return None
        s = str(v).strip().lstrip('#')
        if len(s) == 3:
            s = ''.join(c * 2 for c in s)
        try:
            return RGBColor.from_string(s.upper())
        except Exception:
            return None

    def set_run_font(run, fam=None, size=None, bold=None, italic=None, color=None, underline=False):
        f = run.font
        if fam:
            f.name = fam
            # also set east-asian / complex-script typefaces
            rPr = run._r.get_or_add_rPr()
            for tag in ('a:latin', 'a:ea', 'a:cs'):
                el = rPr.find(qn(tag))
                if el is None:
                    el = rPr.makeelement(qn(tag), {})
                    # order: latin, ea, cs after fill/effects... just append is ok
                    el.set('typeface', fam)
                    rPr.append(el)
                else:
                    el.set('typeface', fam)
        if size is not None:
            f.size = Pt(size)
        if bold is not None:
            f.bold = bold
        if italic is not None:
            f.italic = italic
        if color is not None:
            c = hex_color(color)
            if c:
                f.color.rgb = c
        if underline:
            f.underline = True

    def style_paragraph(p, ps, default_align=None):
        align = ps.get('align', default_align)
        if align:
            m = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT,
                 'justify': PP_ALIGN.JUSTIFY}
            p.alignment = m.get(align, None) if isinstance(align, str) else align
        if ps.get('level'):
            p.level = int(ps['level'])
        if ps.get('spaceBefore'):
            p.space_before = Pt(ps['spaceBefore'])
        if ps.get('spaceAfter'):
            p.space_after = Pt(ps['spaceAfter'])

    def fill_textbox(tb, text_or_paras, defaults):
        tf = tb.text_frame if hasattr(tb, 'text_frame') else tb
        tf.word_wrap = True
        defaults = dict(defaults or {})
        paras = text_or_paras
        if isinstance(paras, str):
            lines = [ln for ln in paras.split('\n')]
            paras = [{'text': ln} for ln in lines]
        first = True
        for ps in paras:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            style_paragraph(p, ps, defaults.get('align'))
            text = ps.get('text', '')
            run = p.add_run()
            run.text = text
            set_run_font(run,
                         ps.get('font', defaults.get('font')),
                         ps.get('size', defaults.get('size')),
                         ps.get('bold', defaults.get('bold')),
                         ps.get('italic', defaults.get('italic')),
                         ps.get('color', defaults.get('color')),
                         ps.get('underline', False))
        try:
            va = defaults.get('valign')
            if va in ('top', 'middle', 'bottom'):
                tf.vertical_anchor = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE,
                                      'bottom': MSO_ANCHOR.BOTTOM}[va]
        except Exception:
            pass
        if defaults.get('lineSpacing'):
            for p in tf.paragraphs:
                p.line_spacing = defaults['lineSpacing']
        if defaults.get('margin'):
            m = defaults['margin']
            for attr in ('margin_left', 'margin_right', 'margin_top', 'margin_bottom'):
                if m.get(attr) is not None:
                    setattr(tf, attr, Inches(m[attr]))

    def add_shape_fill(shp, fill_spec):
        if not fill_spec:
            shp.fill.background()
            return
        if isinstance(fill_spec, str):
            fill_spec = {'color': fill_spec}
        if 'color' in fill_spec:
            c = hex_color(fill_spec['color'])
            if c:
                shp.fill.solid()
                shp.fill.fore_color.rgb = c
            else:
                shp.fill.background()
        elif 'none' in fill_spec:
            shp.fill.background()
        else:
            shp.fill.background()

    def add_line(shp, line_spec):
        if not line_spec:
            try:
                shp.line.fill.background()
            except Exception:
                pass
            return
        c = hex_color(line_spec.get('color', '000000') if isinstance(line_spec, dict) else line_spec)
        if c:
            shp.line.color.rgb = c
        if isinstance(line_spec, dict) and line_spec.get('widthPt'):
            shp.line.width = Pt(line_spec['widthPt'])

    def set_bg(slide, bg):
        if not bg:
            return
        if isinstance(bg, str):
            bg = {'type': bg}
        bt = bg.get('type', 'solid')
        if bt == 'solid' or (bt in ('', 'color') and bg.get('color')):
            c = hex_color(bg.get('color', 'FFFFFF'))
            if c:
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = c
        elif bt == 'none':
            try:
                slide.background.fill.background()
            except Exception:
                pass
        elif bt == 'image':
            img = bg.get('path')
            if img and os.path.exists(img):
                try:
                    img_part, rId = slide.part.get_or_add_image_part(img)
                except Exception as e:
                    warnings.append('bg image failed: %s' % e)
                    return
                cSld = slide._element.find(qn('p:cSld'))
                if cSld is not None:
                    # remove existing bg
                    for old in cSld.findall(qn('p:bg')):
                        cSld.remove(old)
                    bgEl = cSld.makeelement(qn('p:bg'), {})
                    bgPr = bgEl.makeelement(qn('p:bgPr'), {})
                    bgPr.set('dpi', '0')
                    bgPr.set('rotWithShape', '1')
                    blipFill = bgPr.makeelement(qn('a:blipFill'), {})
                    blip = blipFill.makeelement(qn('a:blip'), {})
                    blip.set(qn('r:embed'), rId)
                    blipFill.append(blip)
                    stretch = blipFill.makeelement(qn('a:stretch'), {})
                    stretch.append(blipFill.makeelement(qn('a:fillRect'), {}))
                    blipFill.append(stretch)
                    bgPr.append(blipFill)
                    eff = bgPr.makeelement(qn('a:effectLst'), {})
                    bgPr.append(eff)
                    bgEl.append(bgPr)
                    # p:bg admits exactly one of p:bgPr | p:bgRef -> only bgPr
                    cSld.insert(0, bgEl)
            else:
                warnings.append('bg image not found: %s' % img)
        elif bt == 'gradient':
            if bg.get('color') is None and not bg.get('stops'):
                return
            try:
                csld = slide._element.find(qn('p:cSld'))
                for old in csld.findall(qn('p:bg')):
                    csld.remove(old)
                bgEl = csld.makeelement(qn('p:bg'), {})
                bgPr = bgEl.makeelement(qn('p:bgPr'), {})
                grad = bgPr.makeelement(qn('a:gradFill'), {})
                grad.set('flip', bg.get('flip', 'none'))
                lst = grad.makeelement(qn('a:gsLst'), {})
                stops = bg.get('stops') or [[0, bg.get('color')], [1, bg.get('color2', bg.get('color'))]]
                for pos, col in stops:
                    g = grad.makeelement(qn('a:gs'), {})
                    g.set('pos', str(int(float(pos) * 100000)))
                    srgb = g.makeelement(qn('a:srgbClr'), {})
                    srgb.set('val', str(col).lstrip('#'))
                    g.append(srgb)
                    lst.append(g)
                grad.append(lst)
                lin = grad.makeelement(qn('a:lin'), {})
                lin.set('ang', str(int(float(bg.get('angle', 90)) * 60000)))
                lin.set('scaled', '1')
                grad.append(lin)
                bgPr.append(grad)
                eff = bgPr.makeelement(qn('a:effectLst'), {})
                bgPr.append(eff)
                bgEl.append(bgPr)
                csld.insert(0, bgEl)
            except Exception as e:
                warnings.append('gradient bg failed: %s' % e)

    def add_text(slide, spec_d, defaults):
        merged = dict(defaults or {})
        merged.update(spec_d)
        x = Inches(merged.get('x', 0.5))
        y = Inches(merged.get('y', 0.5))
        w = Inches(merged.get('w', SLIDE_W - 1.0))
        h = Inches(merged.get('h', 1.0))
        tb = slide.shapes.add_textbox(x, y, w, h)
        fill_textbox(tb, merged.get('paragraphs', merged.get('text', '')), merged)
        return tb

    def add_rect(slide, spec_d, rounded=True):
        shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
        shp = slide.shapes.add_shape(shp_type, Inches(spec_d['x']), Inches(spec_d['y']),
                                     Inches(spec_d.get('w', 1)), Inches(spec_d.get('h', 0.5)))
        add_shape_fill(shp, spec_d.get('fill', {'color': spec_d.get('color', 'EEEEEE')}))
        add_line(shp, spec_d.get('line'))
        if spec_d.get('text') or spec_d.get('paragraphs'):
            fill_textbox(shp.text_frame, spec_d.get('paragraphs', spec_d.get('text')), spec_d)
        if rounded and spec_d.get('radius'):
            try:
                shp.adjustments[0] = float(spec_d['radius'])
            except Exception:
                pass
        return shp

    def add_ellipse(slide, spec_d):
        shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(spec_d['x']), Inches(spec_d['y']),
                                     Inches(spec_d.get('w', 1)), Inches(spec_d.get('h', 1)))
        add_shape_fill(shp, spec_d.get('fill', {'color': spec_d.get('color', 'EEEEEE')}))
        add_line(shp, spec_d.get('line'))
        if spec_d.get('text') or spec_d.get('paragraphs'):
            fill_textbox(shp.text_frame, spec_d.get('paragraphs', spec_d.get('text')), spec_d)
        return shp

    def add_line_shape(slide, spec_d):
        from pptx.enum.shapes import MSO_CONNECTOR
        conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                          Inches(spec_d['x1']), Inches(spec_d['y1']),
                                          Inches(spec_d['x2']), Inches(spec_d['y2']))
        c = hex_color(spec_d.get('color', '888888'))
        if c:
            conn.line.color.rgb = c
        conn.line.width = Pt(spec_d.get('widthPt', 1))
        return conn

    def add_table(slide, spec_d):
        rows = spec_d.get('rows', [])
        if not rows:
            return None
        ncols = max(len(r) for r in rows)
        nrows = len(rows)
        gf = slide.shapes.add_table(nrows, ncols, Inches(spec_d['x']), Inches(spec_d['y']),
                                    Inches(spec_d.get('w', 8)), Inches(spec_d.get('h', 0.4 * nrows)))
        tbl = gf.table
        widths = spec_d.get('colWidths')
        if widths:
            for i, wv in enumerate(widths[:ncols]):
                tbl.columns[i].width = Inches(wv)
        header = bool(spec_d.get('header', True))
        hcolor = spec_d.get('headerColor', '1F4E79')
        for ri, row in enumerate(rows):
            for ci in range(ncols):
                cell = tbl.cell(ri, ci)
                val = row[ci] if ci < len(row) else ''
                cell.text_frame.clear()
                p = cell.text_frame.paragraphs[0]
                run = p.add_run()
                run.text = str(val)
                ish = header and ri == 0
                set_run_font(run, spec_d.get('font', '微软雅黑'),
                             spec_d.get('size', 12) if not ish else spec_d.get('headerSize', spec_d.get('size', 12)),
                             ish or spec_d.get('bold', False))
                if ish:
                    try:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = hex_color(hcolor, '1F4E79')
                        run.font.color.rgb = hex_color(spec_d.get('headerTextColor', 'FFFFFF'))
                    except Exception:
                        pass
        return tbl

    slides_spec = spec.get('slides', [])
    for si, sspec in enumerate(slides_spec):
        slide = prs.slides.add_slide(blank)
        # background
        set_bg(slide, sspec.get('bg'))
        dflt = {
            'font': spec.get('theme', {}).get('fontMinor') or '微软雅黑',
            'color': '333333',
            'align': sspec.get('align', 'left'),
            'x': 0.6, 'y': 0.5, 'w': SLIDE_W - 1.2,
        }
        # convenience title/subtitle/bullets
        if sspec.get('title'):
            tspec = sspec.get('title', {})
            if isinstance(tspec, str):
                tspec = {'text': tspec}
            tspec.setdefault('x', 0.6)
            tspec.setdefault('y', 0.4)
            tspec.setdefault('w', SLIDE_W - 1.2)
            tspec.setdefault('h', 1.0)
            tspec.setdefault('size', 40)
            tspec.setdefault('bold', True)
            tspec.setdefault('color', sspec.get('accentColor', '1F4E79') or '1F4E79')
            tspec.setdefault('align', 'center')
            if not tspec.get('text') and not tspec.get('paragraphs'):
                tspec['text'] = sspec.get('title') if isinstance(sspec.get('title'), str) else ''
            add_text(slide, tspec, dflt)
            dflt['y'] = 1.7
        if sspec.get('subtitle'):
            sub_src = sspec['subtitle']
            if isinstance(sub_src, str):
                sub = {'text': sub_src, 'size': 20, 'color': '666666', 'align': 'center',
                       'x': 0.6, 'y': 1.45, 'w': SLIDE_W - 1.2, 'h': 0.6}
            else:
                sub = dict(sub_src)
                sub.setdefault('size', 20)
                sub.setdefault('color', '666666')
                sub.setdefault('align', 'center')
                sub.setdefault('x', 0.6)
                sub.setdefault('y', 1.45)
                sub.setdefault('w', SLIDE_W - 1.2)
                sub.setdefault('h', 0.6)
            add_text(slide, sub, dflt)
            dflt['y'] = 2.3
        if sspec.get('bullets'):
            bs = []
            for bl in sspec['bullets']:
                bs.append({'text': '• ' + bl, 'size': 20, 'spaceAfter': 8})
            add_text(slide, {'paragraphs': bs, 'x': 0.9, 'y': dflt['y'], 'w': SLIDE_W - 1.6,
                             'h': SLIDE_H - dflt['y'] - 0.4, 'size': 20, 'lineSpacing': 1.1},
                     dflt)
        for shp_spec in sspec.get('shapes', []):
            kind = shp_spec.get('kind', 'text')
            eff = {**dflt, **shp_spec}  # inherit slide defaults (font/color/align)
            try:
                if kind == 'text':
                    add_text(slide, eff, dflt)
                elif kind == 'rect':
                    add_rect(slide, eff, bool(shp_spec.get('rounded', True)))
                elif kind == 'ellipse':
                    add_ellipse(slide, eff)
                elif kind == 'line':
                    add_line_shape(slide, shp_spec)
                elif kind == 'image':
                    img = shp_spec.get('path')
                    if img and os.path.exists(img):
                        slide.shapes.add_picture(img, Inches(shp_spec.get('x', 0)), Inches(shp_spec.get('y', 0)),
                                                 Inches(shp_spec.get('w', 4)) if shp_spec.get('w') is not None else None,
                                                 Inches(shp_spec.get('h', 3)) if shp_spec.get('h') is not None else None)
                    else:
                        warnings.append('image not found: %s (slide %d)' % (img, si + 1))
                elif kind == 'table':
                    add_table(slide, eff)
                else:
                    warnings.append('unknown shape kind: %s (slide %d)' % (kind, si + 1))
            except Exception as e:
                warnings.append('shape error slide %d kind %s: %s' % (si + 1, kind, e))

        # ---- notes
        if sspec.get('notes'):
            try:
                slide.notes_slide.notes_text_frame.text = sspec['notes']
            except Exception:
                pass

        # ---- transition
        tr = sspec.get('transition')
        if tr:
            add_transition(slide, tr)
        # ---- animations
        anims = sspec.get('animations')
        if anims:
            add_animations(slide, anims, warnings)

    # ---- theme (fonts + accent colors)
    apply_theme(prs, spec, warnings)

    # ---- show settings
    apply_show(prs, spec)

    # ---- save
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or '.', exist_ok=True)
    prs.save(out_path)
    return {'out': out_path, 'slideCount': len(slides_spec), 'warnings': warnings}

def _insert_sld_child(p_sld, qn, elm):
    """Insert an element that belongs after p:cSld (transition/timing), respecting
    schema order: p:cSld, p:clrMapOvr, p:transition, p:timing."""
    clrMapOvr = p_sld.find(qn('p:clrMapOvr'))
    cSld = p_sld.find(qn('p:cSld'))
    anchor = clrMapOvr if clrMapOvr is not None else cSld
    if anchor is not None:
        anchor.addnext(elm)
    else:
        p_sld.append(elm)

# -- transition XML ----------------------------------------------------------
def add_transition(slide, tr):
    from pptx.oxml.ns import qn
    if isinstance(tr, str):
        tr = {'type': tr}
    ttype = tr.get('type', 'fade')
    p_sld = slide._element
    old = p_sld.find(qn('p:transition'))
    if old is not None:
        p_sld.remove(old)
    if ttype in ('none', False):
        return
    trans = p_sld.makeelement(qn('p:transition'), {})
    spd = tr.get('spd', 'med')
    if spd in ('slow', 'med', 'fast'):
        trans.set('spd', spd)
    else:
        trans.set('spd', 'med')
    if tr.get('advanceAfterSec') is not None:
        trans.set('advTm', str(int(float(tr['advanceAfterSec']) * 1000)))
        trans.set('advClick', '0' if tr.get('advanceOnClick') is False else '1')
    else:
        trans.set('advClick', '0' if tr.get('advanceOnClick') is False else '1')
    effect_map = {
        'fade': 'p:fade', 'wipe': 'p:wipe', 'push': 'p:push', 'cover': 'p:cover',
        'split': 'p:split', 'zoom': 'p:zoom', 'dissolve': 'p:dissolve', 'wheel': 'p:wheel',
        'newsflash': 'p:newsflash', 'random': 'p:random',
    }
    tag = effect_map.get(ttype)
    if tag is None:
        tag = 'p:fade'
    eff = p_sld.makeelement(qn(tag), {})
    if ttype in ('wipe', 'push', 'cover'):
        eff.set('dir', tr.get('dir', 'l'))
    if ttype == 'split':
        eff.set('orient', tr.get('orient', 'vert'))
        eff.set('dir', tr.get('dir', 'l'))
    if ttype == 'zoom':
        eff.set('dir', tr.get('dir', 'in'))
    trans.append(eff)
    _insert_sld_child(p_sld, qn, trans)

# -- animations XML ----------------------------------------------------------
PRESET = {
    'appear':  (1,   'entr', 0),
    'fly':     (2,   'entr', 4),   # from bottom
    'fade':    (10,  'entr', 0),
    'dissolve': (9,  'entr', 0),
    'zoom':    (23,  'entr', 16),
    'wipe':    (22,  'entr', 4),   # from bottom
    'spin':    (46,  'entr', 4),
    'floatUp': (46,  'entr', 4),
}

def add_animations(slide, anims, warnings):
    from pptx.oxml.ns import qn
    p_sld = slide._element
    old = p_sld.find(qn('p:timing'))
    if old is not None:
        p_sld.remove(old)
    # collect ids in document order of sp/pic/graphicFrame/cxnSp (use p:cNvPr@id)
    spTree = p_sld.find('.//' + qn('p:spTree'))
    ids = []
    idnames = {}
    if spTree is not None:
        for s in spTree:
            if local(s.tag) in ('sp', 'pic', 'graphicFrame', 'cxnSp'):
                cNvPr = s.find('.//' + qn('p:cNvPr'))
                if cNvPr is not None:
                    sid = cNvPr.get('id')
                    ids.append(sid)
                    idnames[cNvPr.get('name')] = sid
    if not ids:
        return
    # resolve targets: int index (order added) or shape name
    resolved = []
    for a in anims:
        tgt = a.get('target')
        spid = None
        if isinstance(tgt, int):
            if 0 <= tgt < len(ids):
                spid = ids[tgt]
        elif isinstance(tgt, str):
            if tgt in idnames:
                spid = idnames[tgt]
        if spid is None:
            warnings.append('animation target not found: %r' % (tgt,))
            continue
        resolved.append((spid, a))

    # ---- build timing tree
    p = qn('p:cTn'); a = qn('a:cTn')
    NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    def el(tag, attrs=None, children=None, text=None):
        e = slide._element.makeelement(tag, attrs or {})
        if text is not None:
            e.text = text
        for c in (children or []):
            e.append(c)
        return e

    def cond_el(delay):
        return el(qn('p:cond'), {'delay': delay})
    def tgtsld_el():
        return el(qn('p:tgtEl'), None, [el(qn('p:sldTgt'))])
    def tgtsp_el(spid):
        return el(qn('p:tgtEl'), None, [el(qn('p:spTgt'), {'spid': spid})])
    def attr_name_el(name):
        return el(qn('p:attrName'), None, None, text=name)

    main_children = []
    for spid, a in resolved:
        effect = a.get('effect', 'appear')
        presetID, presetClass, subtype = PRESET.get(effect, PRESET['appear'])
        delay_ms = int(float(a.get('delaySec', 0.3)) * 1000)
        dur_ms = int(float(a.get('durationSec', 0.5)) * 1000)
        trigger = a.get('trigger', 'click')  # click | afterPrev | withPrev
        nodeType = {'click': 'clickEffect', 'afterPrev': 'afterEffect', 'withPrev': 'withEffect'}[trigger] if trigger in ('click', 'afterPrev', 'withPrev') else 'clickEffect'

        def set_visibility_behavior():
            return el(qn('p:set'), None, [
                el(qn('p:cBhvr'), None, [
                    el(qn('p:cTn'), {'id': '7', 'dur': '1', 'fill': 'hold'}, None),
                    tgtsp_el(spid),
                    el(qn('p:attrNameLst'), None, [attr_name_el('style.visibility')]),
                ]),
                el(qn('p:to'), None, [el(qn('p:strVal'), {'val': 'visible'})]),
            ])

        def anim_effect_behavior(filter_name):
            return el(qn('p:animEffect'), {'transition': 'in', 'filter': filter_name}, [
                el(qn('p:cBhvr'), None, [
                    el(qn('p:cTn'), {'id': '7', 'dur': str(dur_ms), 'fill': 'hold'}, None),
                    tgtsp_el(spid),
                    el(qn('p:attrNameLst'), None, [attr_name_el('style.visibility')]),
                ]),
            ])

        def motion_behavior():
            return el(qn('p:anim'), None, [
                el(qn('p:cBhvr'), None, [
                    el(qn('p:cTn'), {'id': '7', 'dur': str(dur_ms), 'fill': 'hold'}, None),
                    tgtsp_el(spid),
                    el(qn('p:attrNameLst'), None, [
                        attr_name_el('ppt_x'),
                        attr_name_el('ppt_y'),
                    ]),
                ]),
                el(qn('p:tavLst'), None, [
                    el(qn('p:tav'), {'tm': '0'}, [el(qn('p:val'), None, [el(qn('p:strVal'), {'val': '#ppt_x'})])]),
                    el(qn('p:tav'), {'tm': '100000'}, [el(qn('p:val'), None, [el(qn('p:strVal'), {'val': '#ppt_x+0.00000'})])]),
                    el(qn('p:tav'), {'tm': '0'}, [el(qn('p:val'), None, [el(qn('p:strVal'), {'val': '#ppt_y'})])]),
                    el(qn('p:tav'), {'tm': '100000'}, [el(qn('p:val'), None, [el(qn('p:strVal'), {'val': '#ppt_y+0.00650'})])]),
                ]),
            ])

        # behavior elements (order inside childTnLst matters: set first)
        behaviors = []
        if effect == 'appear':
            behaviors = [set_visibility_behavior()]
            presetID = 1
        elif effect in ('fade', 'dissolve', 'zoom'):
            if effect == 'dissolve':
                behaviors = [set_visibility_behavior(), anim_effect_behavior('dissolve')]
            else:
                behaviors = [set_visibility_behavior(), anim_effect_behavior('fade')]
        else:  # fly / wipe / spin -> position animation
            behaviors = [set_visibility_behavior(), motion_behavior()]
            presetID, presetClass, subtype = (2, 'entr', 4) if effect == 'fly' else (22, 'entr', 4)
        # click node cTn
        behavior_body = el(qn('p:par'), None, [
            el(qn('p:cTn'), {'id': '5', 'presetID': str(presetID), 'presetClass': presetClass,
                             'presetSubtype': str(subtype), 'fill': 'hold', 'grpId': '0',
                             'nodeType': nodeType}, [
                el(qn('p:stCondLst'), None, [cond_el(str(delay_ms))]),
                el(qn('p:childTnLst'), None, behaviors),
            ]),
        ])
        # withEffect/afterEffect wrap another level? Keep flat: click node only.
        main_children.append(behavior_body)

    seq = el(qn('p:seq'), {'concurrent': '1', 'nextAc': 'seek'}, [
        el(qn('p:cTn'), {'id': '2', 'dur': 'indefinite', 'nodeType': 'mainSeq'}, [
            el(qn('p:childTnLst'), None, main_children),
        ]),
        el(qn('p:prevCondLst'), None, [
            el(qn('p:cond'), {'evt': 'onPrev', 'delay': '0'}, [tgtsld_el()]),
        ]),
        el(qn('p:nextCondLst'), None, [
            el(qn('p:cond'), {'evt': 'onNext', 'delay': '0'}, [tgtsld_el()]),
        ]),
    ])
    root = el(qn('p:par'), None, [
        el(qn('p:cTn'), {'id': '1', 'dur': 'indefinite', 'restart': 'never', 'nodeType': 'tmRoot'}, [
            el(qn('p:childTnLst'), None, [seq]),
        ]),
    ])
    timing = el(qn('p:timing'), None, [el(qn('p:tnLst'), None, [root])])
    # order: cSld, clrMapOvr, transition, timing
    anchor = p_sld.find(qn('p:transition'))
    if anchor is None:
        anchor = p_sld.find(qn('p:clrMapOvr'))
    if anchor is None:
        anchor = p_sld.find('.//' + qn('p:cSld'))
    if anchor is not None:
        anchor.addnext(timing)
    else:
        p_sld.append(timing)

# -- theme -------------------------------------------------------------------
def apply_theme(prs, spec, warnings):
    theme_spec = spec.get('theme', {}) or {}
    fonts = theme_spec.get('fonts', {}) or {}
    colors = theme_spec.get('colors', {}) or {}
    if not fonts and not colors:
        return
    # locate theme part
    themepart = None
    for part in prs.part.package.iter_parts():
        if str(part.partname).startswith('/ppt/theme/'):
            themepart = part
            break
    if themepart is None:
        return
    import lxml.etree as lxmlET
    root = lxmlET.fromstring(themepart.blob)
    M = {'a': A_NS}
    def set_text(mapping, key, val):
        major = root.find('.//a:themeElements/a:fontScheme/a:majorFont/a:latin', M)
        minor = root.find('.//a:themeElements/a:fontScheme/a:minorFont/a:latin', M)
        if mapping == 'major' and major is not None:
            major.set('typeface', val)
        if mapping == 'minor' and minor is not None:
            minor.set('typeface', val)
        for scope, lat in (('major', major), ('minor', minor)):
            if mapping == scope and lat is not None:
                ea = lat.getparent().find('a:ea', M)
                if ea is None:
                    ea = lat.getparent().makeelement('{%s}ea' % A_NS, {})
                    lat.getparent().append(ea)
                ea.set('typeface', val)
    if isinstance(fonts, dict):
        if fonts.get('major'):
            set_text('major', None, fonts['major'])
        if fonts.get('minor'):
            set_text('minor', None, fonts['minor'])
    elif isinstance(fonts, str):
        set_text('major', None, fonts)
        set_text('minor', None, fonts)
    if colors:
        scheme = root.find('.//a:themeElements/a:clrScheme', M)
        if scheme is not None:
            for cname, cval in colors.items():
                found = None
                for c in scheme:
                    if local(c.tag) == cname:
                        found = c
                        break
                if found is None:
                    continue
                v = str(cval).lstrip('#').upper()
                for child in list(found):
                    found.remove(child)
                srgb = found.makeelement('{%s}srgbClr' % A_NS, {'val': v})
                found.append(srgb)
    themepart._blob = lxmlET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

# -- show settings ------------------------------------------------------------
def apply_show(prs, spec):
    show = spec.get('show', {}) or {}
    if not show:
        return
    from pptx.oxml.ns import qn
    presEl = prs._element
    old = presEl.find(qn('p:showPr'))
    if old is not None:
        presEl.remove(old)
    showPr = presEl.makeelement(qn('p:showPr'), {})
    if show.get('loop'):
        showPr.set('loop', '1')
    showPr.set('showAnimation', '0' if show.get('showAnimation') is False else '1')
    showPr.set('showNarration', '0' if show.get('showNarration') is False else '1')
    showPr.set('useTimings', '0' if show.get('useTimings') is False else '1')
    st = show.get('show', 'all')
    if st == 'custom' and show.get('customShow'):
        cust = presEl.makeelement(qn('p:custShow'), {})
        cust.set('st', str(show.get('first', 1)))
        showPr.append(cust)
    elif st == 'range' and show.get('last'):
        rg = presEl.makeelement(qn('p:sldRg'), {})
        rg.set('st', str(show.get('first', 1)))
        rg.set('end', str(show.get('last')))
        showPr.append(rg)
    elif st == 'sldNum' and show.get('first'):
        nbr = presEl.makeelement(qn('p:sldNbr'), {})
        nbr.set('st', str(show.get('first', 1)))
        showPr.append(nbr)
    else:
        allEl = presEl.makeelement(qn('p:sldAll'), {})
        showPr.append(allEl)
    # place showPr right after sldSz / before defaultTextStyle (schema order)
    sldSz = presEl.find(qn('p:sldSz'))
    if sldSz is not None:
        sldSz.addnext(showPr)
    else:
        presEl.append(showPr)

# ---------------------------------------------------------------------------
def demo(out_path):
    spec = {
        'slideWidth': 13.333, 'slideHeight': 7.5,
        'theme': {'fonts': {'major': '微软雅黑', 'minor': '微软雅黑'},
                  'colors': {'accent1': '1F4E79'}},
        'show': {'loop': False, 'show': 'all', 'useTimings': True},
        'slides': [
            {'title': 'PPTX 插件演示', 'subtitle': '读取 + 制作 + 实时展示（MS PowerPoint）',
             'bg': {'type': 'solid', 'color': 'F2F7FB'},
             'transition': {'type': 'fade'},
             'animations': [{'target': 0, 'effect': 'fly', 'trigger': 'afterPrev'}],
             'notes': '这是测试幻灯片'},
            {'bullets': ['可以读取文字/字体/颜色/背景', '可以读取过渡与动画', '可以调用 PowerPoint 边做边看'],
             'bg': {'type': 'gradient', 'stops': [[0, '1F4E79'], [1, '3B8CC4']], 'angle': 90},
             'transition': {'type': 'wipe', 'dir': 'l'},
             'animations': [{'target': 0, 'effect': 'fade', 'trigger': 'afterPrev'}]},
        ],
    }
    return build(spec, out_path)

# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    p1 = sub.add_parser('inspect')
    p1.add_argument('file')
    p2 = sub.add_parser('build')
    p2.add_argument('spec')
    p2.add_argument('out')
    p2.add_argument('--to-json', action='store_true')
    p3 = sub.add_parser('demo')
    p3.add_argument('out')
    args = ap.parse_args()
    if args.cmd == 'inspect':
        try:
            print(json.dumps(inspect(args.file), ensure_ascii=False, indent=1))
        except Exception as e:
            print(json.dumps({'error': str(e), 'trace': __import__('traceback').format_exc()},
                             ensure_ascii=False), file=sys.stderr)
            sys.exit(2)
    elif args.cmd == 'build':
        with open(args.spec, 'r', encoding='utf8') as f:
            spec = json.load(f)
        res = build(spec, args.out)
        print(json.dumps(res, ensure_ascii=False))
    elif args.cmd == 'demo':
        print(json.dumps(demo(args.out), ensure_ascii=False))

if __name__ == '__main__':
    main()
