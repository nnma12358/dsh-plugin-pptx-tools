# pptx spec 字段参考

`pptx_build` 接受的就是这份 JSON。字段全部可选，除非注明必填。
尺寸单位一律是**英寸**，字号是**磅（pt）**，颜色是**不带 `#` 的十六进制字符串**。

可从 `examples/` 直接抄：`v5.json`（最小）、`three.json`、`nodejs-talk-spec.json`（12 页真实案例，含代码块/表格/渐变）。

---

## 1. 顶层

| 字段 | 类型 | 默认 | 说明 |
|---|---|---|---|
| `slideWidth` | number | `13.333` | 页面宽（16:9） |
| `slideHeight` | number | `7.5` | 页面高 |
| `theme` | object | — | 只有给了 `theme.fonts` 或 `theme.colors` 才处理 |
| `theme.fonts` | `{major, minor}` 或 string | — | 同时写 `a:latin` 与 `a:ea`（中文字体靠 `ea`） |
| `theme.colors` | `{方案色名: 十六进制}` | — | 只替换主题里**已存在**的 `a:clrScheme` 名称，未知名称静默跳过 |
| `theme.fontMinor` | string | — | 未文档化，作用于形状默认字体 |
| `show` | object | — | 放映设置，见下 |
| `slides` | array | `[]` | 幻灯片数组；`slideCount` 即其长度 |

### `show`

| 字段 | 默认 | 说明 |
|---|---|---|
| `loop` | false | 循环放映 |
| `showAnimation` | true | 传 `false` 才写 `0` |
| `showNarration` | true | 同上 |
| `useTimings` | true | 同上 |
| `show` | — | `custom` \| `range` \| `sldNum` \| 其它→`all` |
| `first` / `last` | `1` / — | `range` 模式需要 |

> 已知不足：`customShow` 只做真值判断、值未被使用，也不会生成 `p:custShowLst`，自定义放映实际不可用。

---

## 2. 每页 `slides[]`

| 字段 | 类型 | 默认 | 说明 |
|---|---|---|---|
| `bg` | string 或 object | — | 见下 |
| `align` | `left`\|`center`\|`right`\|`justify` | `left` | 本页形状的默认对齐 |
| `accentColor` | 十六进制 | `1F4E79` | 标题默认色 |
| `title` | string 或 object | — | 字符串时自动套默认版式 |
| `subtitle` | string 或 object | — | 同上 |
| `bullets` | string[] | — | 每项自动加 `• ` 前缀 |
| `shapes` | array | — | 见第 3 节 |
| `notes` | string | — | 备注页文字 |
| `transition` | string 或 object | — | 见第 4 节 |
| `animations` | array | — | 见第 5 节 |

`title` 省略时的默认版式：x 0.6, y 0.4, w `slideWidth-1.2`, h 1.0, size **40**, bold, 居中。
`subtitle`：size **20**, 色 `666666`, 居中, x 0.6, y 1.45, h 0.6。
`bullets`：size 20, spaceAfter 8, x 0.9, w `slideWidth-1.6`, lineSpacing 1.1。

### `bg`

| `type` | 需要的字段 | 说明 |
|---|---|---|
| `solid`（默认） | `color`（默认 `FFFFFF`） | 纯色 |
| `none` | — | 无背景 |
| `image` | `path` | 文件不存在只产生 warning `bg image not found: …` |
| `gradient` | `stops: [[位置, 色], …]` | 缺省回退成 `[[0,color],[1,color2]]` |

渐变还可用 `angle`（度，默认 **90**）与 `flip`（默认 `none`）。

---

## 3. `shapes[]`

`kind` 默认 `text`；取值 `text` / `rect` / `ellipse` / `line` / `image` / `table`，
未知取值只产生 warning `unknown shape kind: …`。

形状字段会与页面默认值合并（`x 0.6, y 0.5, w slideWidth-1.2, color 333333`），
**例外：`kind: "line"` 不参与合并**。

`x`、`y` 对除 `line` 外的所有类型都是必填；缺了会 KeyError。

| kind | 必填 | 其它字段与默认 |
|---|---|---|
| `text` | `x`,`y` | `text`（字符串，按 `\n` 拆段）或 `paragraphs`；`w` 默认 `slideWidth-1.0`、`h` 默认 1 |
| `rect` | `x`,`y` | `w` 1、`h` 0.5、`fill` 默认 `{color: 333333 或 EEEEEE}`、`rounded` 默认 **true**、`radius` 映射到 `adjustments[0]` |
| `ellipse` | `x`,`y` | `w`/`h` 默认 1 |
| `line` | `x1`,`y1`,`x2`,`y2` | `color` 默认 `888888`、`widthPt` 默认 1；画成直连接符 |
| `image` | `path`,`x`,`y` | `w` 默认 4、`h` 默认 3；文件不存在产生 warning `image not found: …` |
| `table` | `rows`（非空） | `colWidths`、`header` 默认 true、`headerColor` `1F4E79`、`headerTextColor` `FFFFFF`、`font` `微软雅黑`、`size` 12、`headerSize` = size、`bold` false、`w` 8、`h` 默认 `0.4 × 行数` |

文本类字段（`text` 形状与表格单元格通用）：
`font`、`size`、`bold`、`italic`、`underline`、`color`、`align`、`level`、
`spaceBefore`、`spaceAfter`、`lineSpacing`、`valign`（`top|middle|bottom`）、
`margin`（`margin_left|margin_right|margin_top|margin_bottom`）。

---

## 4. `transition`

| 字段 | 默认 | 取值 |
|---|---|---|
| `type` | `fade` | `fade` `wipe` `push` `cover` `split` `zoom` `dissolve` `wheel` `newsflash` `random`；`none` 或 `false` 表示删除切换 |
| `spd` | `med` | `slow` `med` `fast` |
| `advanceAfterSec` | — | 秒，转成 `advTm` 毫秒 |
| `advanceOnClick` | true | 传 `false` 才写 `0` |
| `dir` | `l` | 用于 `wipe`/`push`/`cover` |
| `orient` | `vert` | 用于 `split` |

---

## 5. `animations[]`

| 字段 | 必填 | 默认 | 说明 |
|---|---|---|---|
| `target` | ✅ | — | 形状在文档顺序里的**下标**（整数）或**形状名**；找不到产生 warning |
| `effect` | | `appear` | `appear` `fly` `fade` `dissolve` `zoom` `wipe` `spin` `floatUp` |
| `trigger` | | `click` | `click` \| `afterPrev` \| `withPrev` |
| `delaySec` | | `0.3` | |
| `durationSec` | | `0.5` | |

> 已知不足：`afterPrev`/`withPrev` 只改节点类型，不会生成嵌套 `p:par` 层级，
> 真正的点击序列行为与 PowerPoint 原生「上一动画之后」不一致；
> 非 fade 效果共用一条写死的微小位移路径，`spin`/`floatUp`/`zoom` 的预设值在该分支实际不可达。

---

## 6. 返回值与错误

- `pptx_build` / `pptx_demo` 返回 `{"out": 路径, "slideCount": N, "warnings": [...]}`。
- **`warnings` 一定要看**：图片缺失、形状类型未知、动画目标找不到都只在这里出现，不会中断生成。
- `pptx_inspect` 失败时退出码 2，并在 stderr 给出 `{"error": …, "trace": …}`。

## 7. 已知无效/易踩的字段

- `pptx_cli.py build --to-json` 是空开关，传了没有效果。
- `make_film_pptx.py` 被当作模块 import 时，`preprocess()` 会 `NameError`（`ImageOps` 只在 `__main__` 分支导入）。
- 两个 film 脚本默认输出**同一个文件名**，连续运行互相覆盖 —— 请显式传 `out`。
- 默认字体是 `微软雅黑`，导出 PNG/查看效果需要目标机器装了这个字体。
