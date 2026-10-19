# dsh-plugin-pptx-tools

PowerPoint 读写与生成。把 `pptx_tools/` 的脚本封装成 6 个 DSH 工具。

## 工具

| 工具 | 说明 |
|---|---|
| `pptx_inspect` | 读 .pptx → JSON 摘要（页面尺寸、形状、文字与字体、背景、切换、动画、备注、主题、媒体） |
| `pptx_build` | 按 JSON spec 生成 .pptx；`spec` 可直接传对象，也可传 `specFile` |
| `pptx_demo` | 生成内置 2 页示例，用于确认环境 |
| `pptx_film_deck` | 目录下照片 → 每张一页的「电影胶卷」，wheel 切换、4 秒自动翻页、循环 |
| `pptx_film_carousel` | 照片 → 单页胶片横向轮播（原生时间轴，每 7 秒滑一个周期） |
| `pptx_powerpoint` | COM 控制本机 PowerPoint：`open`/`show`/`next`/`prev`/`end`/`export`/`close`/`refresh`/`info` |

## 依赖

```bash
python3 -m pip install -r payload/requirements.txt
# python-pptx==1.0.2  Pillow  lxml  XlsxWriter  typing_extensions
```

解释器查找顺序：`DSH_PYTHON` → `python3` → `python` → `py`。
找不到可用解释器时，工具会直接返回该装的 pip 命令，而不是晦涩的报错。

`pptx_powerpoint` 额外需要 Windows + 桌面版 Microsoft PowerPoint；
在 WSL 里通过 interop 调用 `powershell.exe` 也可以（本包已实测通过）。

## 用法示例

```jsonc
// pptx_build 的 spec
{
  "slideWidth": 13.333, "slideHeight": 7.5,
  "theme": { "colors": { "accent1": "2E74B5" },
             "fonts": { "major": "微软雅黑", "minor": "微软雅黑" } },
  "slides": [
    { "title": "标题", "subtitle": "副标题",
      "bg": { "type": "gradient", "stops": [[0, "0B1220"], [1, "1F4E79"]], "angle": 45 },
      "bullets": ["第一条", "第二条"],
      "transition": { "type": "fade", "advanceAfterSec": 3 },
      "shapes": [{ "kind": "rect", "x": 0.8, "y": 4.2, "w": 4, "h": 0.9,
                   "fill": { "color": "F2C14E" } }] }
  ]
}
```

字段全表见 `docs/SPEC.md`；可运行的真实案例见 `examples/nodejs-talk-spec.json`。

## 注意

- **一定看返回里的 `warnings`**：图片缺失、形状类型拼错、动画目标找不到都只出现在这里，不会中断生成。
- `pptx_film_deck` 与 `pptx_film_carousel` 默认输出同名文件，连续运行会互相覆盖 —— 请显式传 `out`。
- 两个 film 脚本会把缩放后的照片写到 `payload/out/film-assets/`，该目录可随时删除重建。
- 默认字体 `微软雅黑` 只是名字字符串；目标机器没装该字体时，PPT 仍能生成，但显示效果会变。
- `validate.py`（COM 端到端自检）随包提供但**未接入工具**，因为它依赖 `C:\Users\Public\Documents\DSH_Pptx\` 这个硬编码暂存目录。

## 安装（dsh desktop）

桌面端 profile 由 Electron 应用独占管理，`dsh plugin --profile desktop add …` 会被官方拒绝
（`profile "desktop" is managed exclusively by the Electron application`）。请任选：

**A. 应用内安装（推荐）**：设置 → 插件 → 添加本地插件，选择本仓库目录。

**B. 手工登记**：
1. 把仓库放到 `%USERPROFILE%\.dsh\profiles\desktop\plugins\<包名>\`
2. 在该 profile 的 `package.json` 里加 `link:` 依赖，并把包名加入 `dsh.profile.bundles`
3. 用桌面版自带 pnpm 装依赖，然后重启 dsh desktop

## 开发

本仓库是纯源码包，无构建步骤：`main` / `exports` 直接指向 `lib/`，
`dsh.bundle.patch` 指向 `cordis.patch.yml`。
