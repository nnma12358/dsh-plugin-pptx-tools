#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate pptx variants and verify each opens cleanly in MS PowerPoint (COM)."""
import json, os, subprocess, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
WIN_DIR = r'C:\Users\Public\Documents\DSH_Pptx\validate'
LIN_DIR = '/mnt/c/Users/Public/Documents/DSH_Pptx/validate'
PS = r'C:\Users\Public\Documents\DSH_Pptx\powerpoint.ps1'
CFG = r'C:\Users\Public\Documents\DSH_Pptx\cfg.json'
os.makedirs(LIN_DIR, exist_ok=True)

def build(name, spec):
    sp = os.path.join(LIN_DIR, name + '.json')
    out = os.path.join(LIN_DIR, name + '.pptx')
    json.dump(spec, open(sp, 'w'), ensure_ascii=False)
    r = subprocess.run([os.path.join(HERE, 'venv/bin/python'), os.path.join(HERE, 'pptx_cli.py'),
                        'build', sp, out], capture_output=True, text=True)
    return r.stdout.strip() + r.stderr.strip()

def com(action, path=None):
    cfg = {'action': action}
    if path:
        cfg['path'] = path
    json.dump(cfg, open('/mnt/c/Users/Public/Documents/DSH_Pptx/cfg.json', 'w'))
    r = subprocess.run(['/mnt/c/WINDOWS/system32/WindowsPowerShell/v1.0/powershell.exe',
                        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', PS, '-Config', CFG],
                       capture_output=True, text=True, encoding='gbk', errors='replace', timeout=180)
    txt = (r.stdout or '').strip() or (r.stderr or '').strip()
    try:
        i = txt.rfind('{')
        return json.loads(txt[i:])
    except Exception:
        return {'ok': False, 'error': txt[-300:]}

def check(name, spec):
    r = build(name, spec)
    print('BUILD %-12s %s' % (name, r))
    com('close')
    res = com('open', os.path.join(WIN_DIR, name + '.pptx'))
    if not res.get('ok'):
        time.sleep(1.5)
        res2 = com('open', os.path.join(WIN_DIR, name + '.pptx'))
        if res2.get('ok'):
            res = res2
    status = 'OK' if res.get('ok') else 'FAIL'
    err = res.get('error', '')
    print('OPEN  %-12s %s  %s' % (name, status, err))
    return res

specs = {}

specs['plain'] = {
    'slideWidth': 13.333, 'slideHeight': 7.5,
    'slides': [{'title': 'Plain 测试', 'subtitle': 'no extras'}],
}

specs['solidbg'] = {
    'slideWidth': 13.333, 'slideHeight': 7.5,
    'slides': [{'title': 'solid', 'bg': {'type': 'solid', 'color': 'F2F7FB'}}],
}

specs['gradient'] = {
    'slideWidth': 13.333, 'slideHeight': 7.5,
    'slides': [{'title': 'gradient', 'bg': {'type': 'gradient', 'stops': [[0, '1F4E79'], [1, '3B8CC4']], 'angle': 90}}],
}

specs['trans'] = {
    'slideWidth': 13.333, 'slideHeight': 7.5,
    'slides': [
        {'title': 'trans', 'transition': {'type': 'fade', 'advanceAfterSec': 3}},
        {'title': 'trans2', 'transition': {'type': 'wipe', 'dir': 'l'}, 'bg': {'type': 'solid', 'color': 'FFFFFF'}},
    ],
}

specs['anim'] = {
    'slideWidth': 13.333, 'slideHeight': 7.5,
    'slides': [{
        'title': 'anim',
        'shapes': [
            {'kind': 'text', 'text': '目标 1', 'x': 1, 'y': 3, 'w': 6, 'h': 1, 'size': 24},
            {'kind': 'text', 'text': '目标 2', 'x': 1, 'y': 4.5, 'w': 6, 'h': 1, 'size': 24},
        ],
        'animations': [
            {'target': 0, 'effect': 'appear', 'trigger': 'click'},
            {'target': 1, 'effect': 'fly', 'trigger': 'afterPrev', 'delaySec': 0.2, 'durationSec': 0.6},
        ],
    }],
}

specs['theme'] = {
    'slideWidth': 13.333, 'slideHeight': 7.5,
    'theme': {'fonts': {'major': '微软雅黑', 'minor': '微软雅黑'}, 'colors': {'accent1': 'C00000'}},
    'slides': [{'title': 'theme'}],
}

specs['show'] = {
    'slideWidth': 13.333, 'slideHeight': 7.5,
    'show': {'loop': True, 'show': 'range', 'first': 1, 'last': 2},
    'slides': [{'title': 'show1'}, {'title': 'show2'}],
}

specs['full'] = {
    'slideWidth': 13.333, 'slideHeight': 7.5,
    'theme': {'fonts': {'major': '微软雅黑', 'minor': '微软雅黑'}, 'colors': {'accent1': 'C00000'}},
    'show': {'loop': False, 'show': 'all'},
    'slides': [
        {'title': '完整测试 1', 'subtitle': 'bg solid + fade + appear',
         'bg': {'type': 'solid', 'color': 'EAF2FA'},
         'transition': {'type': 'fade', 'advanceAfterSec': 4},
         'animations': [{'target': 0, 'effect': 'appear', 'trigger': 'click'}],
         'shapes': [
             {'kind': 'rect', 'x': 0.5, 'y': 1.6, 'w': 3, 'h': 1, 'fill': {'color': 'C00000'}, 'text': '红块', 'size': 18, 'color': 'FFFFFF'},
             {'kind': 'table', 'x': 4, 'y': 5, 'w': 8, 'h': 1.6,
              'rows': [['名称', '数量'], ['苹果', '3'], ['香蕉', '5']]},
         ],
         'animations': [
             {'target': 0, 'effect': 'appear', 'trigger': 'click'},
             {'target': 1, 'effect': 'fly', 'trigger': 'afterPrev'},
         ],
         'notes': '完整测试备注'},
        {'title': '完整测试 2', 'bg': {'type': 'gradient', 'stops': [[0, '1F4E79'], [1, '3B8CC4']], 'angle': 90},
         'transition': {'type': 'wipe', 'dir': 'l'},
         'animations': [{'target': 0, 'effect': 'fade', 'trigger': 'afterPrev'}]},
    ],
}

if __name__ == '__main__':
    only = sys.argv[1:] or list(specs)
    failures = []
    for name in only:
        res = check(name, specs[name])
        if not res.get('ok'):
            failures.append((name, res.get('error', '')))
    print('\n== RESULT:', 'ALL OK' if not failures else ('FAILED: %s' % failures))
