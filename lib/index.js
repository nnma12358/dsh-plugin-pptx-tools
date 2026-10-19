/* ============================================================================
 * dsh-plugin-pptx-tools — PowerPoint (.pptx) 读写与生成（profile bundle, Host half）
 *
 * 把工作区里 pptx_tools/ 的自制脚本封装成 DSH 工具：
 *   pptx_cli.py            → pptx_inspect / pptx_build / pptx_demo
 *   make_film_pptx.py      → pptx_film_deck      每张照片一页的「电影胶卷」
 *   make_film_carousel.py  → pptx_film_carousel  单页胶卷横向轮播
 *   powerpoint.ps1         → pptx_powerpoint     COM 控制本机 PowerPoint
 *
 * pptx_inspect / pptx_build / pptx_film_* 是纯 Python（python-pptx + Pillow + lxml），
 * 跨平台，不需要安装 Office。只有 pptx_powerpoint 需要 Windows + 桌面版 PowerPoint。
 *
 * 依赖：payload/requirements.txt（python-pptx、Pillow、lxml、XlsxWriter）。
 * 解释器顺序：DSH_PYTHON 环境变量 → python3 → python → py。
 * ========================================================================== */

import { spawn } from 'node:child_process'
import { existsSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'pptx-tools'
export const inject = ['tools']

const PAYLOAD = join(dirname(fileURLToPath(import.meta.url)), '..', 'payload')

/* -------------------------------------------------------------------- 进程 */

function run(argv, opts = {}) {
  return new Promise((resolve) => {
    let child
    try {
      child = spawn(argv[0], argv.slice(1), {
        cwd: opts.cwd || PAYLOAD,
        env: { ...process.env, ...(opts.env || {}) },
        windowsHide: true,
      })
    } catch (error) {
      resolve({ code: -1, stdout: '', stderr: String((error && error.message) || error), timedOut: false })
      return
    }
    const cap = opts.maxBytes || 4000000
    let stdout = ''
    let stderr = ''
    let timedOut = false
    let timer = null
    if (opts.timeoutMs) {
      timer = setTimeout(() => {
        timedOut = true
        try { child.kill('SIGKILL') } catch { /* 已经退出 */ }
      }, opts.timeoutMs)
    }
    child.stdout.on('data', (d) => { if (stdout.length < cap) stdout += d.toString('utf8') })
    child.stderr.on('data', (d) => { if (stderr.length < cap) stderr += d.toString('utf8') })
    child.on('error', (error) => {
      if (timer) clearTimeout(timer)
      resolve({ code: -1, stdout, stderr: stderr + String((error && error.message) || error), timedOut })
    })
    child.on('close', (code) => {
      if (timer) clearTimeout(timer)
      resolve({ code, stdout, stderr, timedOut })
    })
  })
}

/* -------------------------------------------------------------- Python 探测 */

const PY_CANDIDATES = ['python3', 'python', 'py']

/* Windows 上 Python 的 stdout 默认走 ANSI 代码页（中文系统是 GBK），
 * 而 inspect 的 JSON 里含 • 等字符，会直接 UnicodeEncodeError 退出。
 * PYTHONUTF8 让解释器进入 UTF-8 模式，PYTHONIOENCODING 兜底非 UTF-8 模式。 */
const PY_ENV = { PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' }
let pythonCache = null

/**
 * dsh desktop 自带一份受管 Python（resources/runtime/primary-runtime/dependencies/python），
 * 并预装 python-pptx / Pillow / lxml —— 桌面版用户机器上通常**没有**系统 Python，
 * 所以必须优先找到它，否则所有 pptx 工具在桌面端都跑不起来。
 *
 * 从进程自身位置向上找，兼容两种宿主：
 *   - Electron 主进程旁：<app>/resources/...
 *   - 内嵌 node.exe：<app>/resources/runtime/primary-runtime/dependencies/node/bin/node.exe
 */
function desktopRuntimePython() {
  const exe = process.platform === 'win32' ? 'python.exe' : 'python3'
  const rel = ['runtime', 'primary-runtime', 'dependencies', 'python', exe]
  const bases = []
  if (process.env.DSH_DESKTOP_RUNTIME) bases.push(process.env.DSH_DESKTOP_RUNTIME)
  if (process.resourcesPath) bases.push(process.resourcesPath)
  let dir = ''
  try { dir = dirname(process.execPath || '') } catch { dir = '' }
  for (let i = 0; i < 7 && dir; i += 1) {
    bases.push(dir, join(dir, 'resources'))
    const parent = dirname(dir)
    if (parent === dir) break
    dir = parent
  }
  for (const base of bases) {
    try {
      const candidate = join(base, ...rel)
      if (existsSync(candidate)) return candidate
    } catch { /* 忽略非法路径 */ }
  }
  return undefined
}

/** 找到第一个「存在且能 import pptx/PIL/lxml」的解释器；结果进程内缓存。 */
async function resolvePython() {
  if (pythonCache) return pythonCache
  const override = process.env.DSH_PYTHON
  const candidates = override ? [override] : [desktopRuntimePython(), ...PY_CANDIDATES].filter(Boolean)
  const probe = 'import pptx, PIL, lxml; print("ok")'
  let firstWorking = null
  for (const candidate of candidates) {
    const r = await run([candidate, '-c', probe], { timeoutMs: 30000, env: PY_ENV })
    if (r.code === 0 && r.stdout.includes('ok')) {
      pythonCache = { command: candidate, ok: true }
      return pythonCache
    }
    if (!firstWorking && r.code !== -1) firstWorking = candidate
  }
  pythonCache = {
    command: firstWorking || candidates[0],
    ok: false,
    detail: '未找到能 import pptx/PIL/lxml 的 Python 解释器。请先安装依赖：\n'
      + '  ' + (firstWorking || 'python3') + ' -m pip install -r "' + join(PAYLOAD, 'requirements.txt') + '"\n'
      + '或把可用的解释器路径写入环境变量 DSH_PYTHON。',
  }
  return pythonCache
}

async function runPy(scriptArgs, opts = {}) {
  const python = await resolvePython()
  if (!python.ok) return { ok: false, text: python.detail }
  const argv = [python.command, join(PAYLOAD, 'pptx_cli.py'), ...scriptArgs]
  const r = await run(argv, { timeoutMs: opts.timeoutMs || 180000, env: { ...PY_ENV, ...(opts.env || {}) } })
  const text = [r.stdout.trim(), r.stderr.trim()].filter(Boolean).join('\n')
  if (r.timedOut) return { ok: false, text: '超时（已终止）。\n' + text }
  if (r.code !== 0) return { ok: false, text: '退出码 ' + String(r.code) + '。\n' + text }
  return { ok: true, text: text || '(无输出)' }
}

/* ------------------------------------------------------------- 辅助与注册 */

function tempFile(suffix) {
  const dir = mkdtempSync(join(tmpdir(), 'dsh-pptx-'))
  return join(dir, 'file' + suffix)
}

function define(ctx, spec) {
  ctx.tools.register({
    name: spec.name,
    description: spec.description,
    parameters: { type: 'object', properties: spec.properties, required: spec.required || [] },
    output: {
      schema: { type: 'string' },
      render: (_args, value) => [{ type: 'text', text: String(value) }],
    },
    execute: async (args) => {
      try {
        return await spec.execute(args || {})
      } catch (error) {
        return '工具内部错误：' + String((error && error.message) || error)
      }
    },
  })
}

/* ------------------------------------------------------------------- 注册 */

export function apply(ctx) {
  /* --- 1. inspect -------------------------------------------------------- */
  define(ctx, {
    name: 'pptx_inspect',
    description: '读取一个 .pptx 并输出 JSON 结构摘要：页面尺寸、每页的形状/文字/字体/颜色、'
      + '背景、切换、动画、备注、主题字体与配色、放映设置、媒体文件。'
      + '修改既有 PPT 前先用它看清现状。',
    properties: {
      file: { type: 'string', description: '.pptx 文件路径（必填）。' },
    },
    required: ['file'],
    async execute(args) {
      if (!args.file) return '缺少必填参数 file。'
      if (!existsSync(args.file)) return '文件不存在：' + args.file
      const r = await runPy(['inspect', args.file], { timeoutMs: 120000 })
      return r.text
    },
  })

  /* --- 2. build ---------------------------------------------------------- */
  define(ctx, {
    name: 'pptx_build',
    description: '按结构化 spec 生成 .pptx。spec 可以直接传对象，也可以传已有的 .json 路径。'
      + '支持的字段：slideWidth/slideHeight（英寸，默认 13.333×7.5）、'
      + 'theme.fonts{major,minor} 与 theme.colors{方案色名:十六进制}、show{loop,...}、'
      + 'slides[]{bg, align, accentColor, title, subtitle, bullets[], shapes[], notes, transition, animations[]}。'
      + 'shapes 的 kind 可为 text/rect/ellipse/line/image/table，x/y/w/h 单位英寸。'
      + 'transition.type 支持 fade/wipe/push/cover/split/zoom/dissolve/wheel/newsflash/random。'
      + 'animations[].effect 支持 appear/fly/fade/dissolve/zoom/wipe/spin/floatUp。'
      + '返回 {out, slideCount, warnings}，warnings 会指出找不到的图片/未知形状类型等。'
      + '完整字段说明见本插件的 docs/SPEC.md。',
    properties: {
      spec: { type: 'object', description: '演示文稿 spec 对象（与 specFile 二选一）。' },
      specFile: { type: 'string', description: 'spec 的 .json 文件路径（与 spec 二选一）。' },
      out: { type: 'string', description: '输出 .pptx 路径（必填）；父目录会自动创建。' },
    },
    required: ['out'],
    async execute(args) {
      if (!args.out) return '缺少必填参数 out。'
      let specPath = args.specFile
      if (!specPath) {
        if (!args.spec || typeof args.spec !== 'object') return '必须提供 spec 对象或 specFile 路径。'
        specPath = tempFile('.json')
        writeFileSync(specPath, JSON.stringify(args.spec), 'utf8')
      } else if (!existsSync(specPath)) {
        return 'spec 文件不存在：' + specPath
      }
      const r = await runPy(['build', specPath, args.out], { timeoutMs: 180000 })
      if (!r.ok) return r.text
      return r.text + '\n输出：' + args.out
    },
  })

  /* --- 3. demo ----------------------------------------------------------- */
  define(ctx, {
    name: 'pptx_demo',
    description: '生成一个 2 页的内置演示文件，用来快速确认 python-pptx 环境是否可用。',
    properties: {
      out: { type: 'string', description: '输出 .pptx 路径（必填）。' },
    },
    required: ['out'],
    async execute(args) {
      if (!args.out) return '缺少必填参数 out。'
      const r = await runPy(['demo', args.out], { timeoutMs: 120000 })
      return r.text
    },
  })

  /* --- 4/5. 电影胶卷 ------------------------------------------------------ */
  const filmProps = {
    srcDir: { type: 'string', description: '照片目录（jpg/png）；不给则用 payload/图片。' },
    out: { type: 'string', description: '输出 .pptx 路径（必填）。' },
  }
  define(ctx, {
    name: 'pptx_film_deck',
    description: '把目录下所有照片做成「电影胶卷」风格：16:9 黑底，每张照片一页，'
      + '上下两条齿孔带，wheel 切换、每 4 秒自动翻页、循环放映。'
      + '照片会等比缩放（长边 ≤2000px，JPEG q88）到 payload/out/film-assets。',
    properties: filmProps,
    required: ['out'],
    async execute(args) {
      if (!args.out) return '缺少必填参数 out。'
      const env = args.srcDir ? { PPTX_SRC_DIR: args.srcDir } : {}
      const python = await resolvePython()
      if (!python.ok) return python.detail
      const r = await run(
        [python.command, join(PAYLOAD, 'make_film_pptx.py'), args.out],
        { timeoutMs: 600000, env: { ...PY_ENV, ...env } },
      )
      const text = [r.stdout.trim(), r.stderr.trim()].filter(Boolean).join('\n')
      if (r.code !== 0) return '退出码 ' + String(r.code) + '。\n' + text
      return text
    },
  })

  define(ctx, {
    name: 'pptx_film_carousel',
    description: '把所有照片排成一条「胶片」放在同一页里，用原生 PowerPoint 时间轴让胶片每 7 秒'
      + '向左滑动一个周期，无限循环——单页轮播，适合当展台/大屏背景。',
    properties: filmProps,
    required: ['out'],
    async execute(args) {
      if (!args.out) return '缺少必填参数 out。'
      const env = args.srcDir ? { PPTX_SRC_DIR: args.srcDir } : {}
      const python = await resolvePython()
      if (!python.ok) return python.detail
      const r = await run(
        [python.command, join(PAYLOAD, 'make_film_carousel.py'), args.out],
        { timeoutMs: 600000, env: { ...PY_ENV, ...env } },
      )
      const text = [r.stdout.trim(), r.stderr.trim()].filter(Boolean).join('\n')
      if (r.code !== 0) return '退出码 ' + String(r.code) + '。\n' + text
      return text
    },
  })

  /* --- 6. PowerPoint COM ------------------------------------------------- */
  define(ctx, {
    name: 'pptx_powerpoint',
    description: '用 COM 控制本机已安装的桌面版 PowerPoint（仅 Windows）：打开、放映、上一页/下一页、'
      + '结束放映、导出每页 PNG、关闭、刷新、查询状态。'
      + 'action 取 open|show|next|prev|end|export|close|refresh|info；'
      + 'path 为 .pptx 的 Windows 路径；export 时用 outdir 指定输出目录，width/height 默认 1280×720。'
      + '注意：这会真的操作本机 PowerPoint 窗口，只在需要放映或导出预览图时使用。',
    properties: {
      action: { type: 'string', description: 'open|show|next|prev|end|export|close|refresh|info（必填）。' },
      path: { type: 'string', description: '.pptx 的 Windows 路径。' },
      slide: { type: 'number', description: 'show 时跳转到第几页。' },
      outdir: { type: 'string', description: 'export 的输出目录（Windows 路径）。' },
      width: { type: 'number', description: 'export 宽度，默认 1280。' },
      height: { type: 'number', description: 'export 高度，默认 720。' },
    },
    required: ['action'],
    async execute(args) {
      if (!args.action) return '缺少必填参数 action。'
      const config = { action: args.action }
      for (const key of ['path', 'slide', 'outdir', 'width', 'height']) {
        if (args[key] !== undefined && args[key] !== null) config[key] = args[key]
      }
      const cfgPath = tempFile('.cfg.json')
      writeFileSync(cfgPath, JSON.stringify(config), 'utf8')

      const ps = process.env.DSH_POWERSHELL
        || (process.platform === 'win32'
          ? 'powershell.exe'
          : (existsSync('/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe')
            ? '/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'
            : 'powershell.exe'))
      let scriptPath = join(PAYLOAD, 'powerpoint.ps1')
      let configHost = cfgPath
      if (process.platform !== 'win32') {
        const conv = await run(['wslpath', '-w', scriptPath], { timeoutMs: 15000 })
        if (conv.code === 0 && conv.stdout.trim()) scriptPath = conv.stdout.trim()
        const conv2 = await run(['wslpath', '-w', cfgPath], { timeoutMs: 15000 })
        if (conv2.code === 0 && conv2.stdout.trim()) configHost = conv2.stdout.trim()
      }
      const r = await run([ps, '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
        scriptPath, '-Config', configHost], { timeoutMs: 300000 })
      const text = [r.stdout.trim(), r.stderr.trim()].filter(Boolean).join('\n')
      if (r.code !== 0) return '退出码 ' + String(r.code) + '（需要 Windows + 桌面版 PowerPoint）。\n' + text
      return text
    },
  })
}
