# wg-screenshot-to-ppt

把截图或图片转换成**可编辑的 PowerPoint（.pptx）**。

普通做法是往幻灯片里贴一张整页图片，结果文字改不了、色块换不了色、箭头挪不动。本项目的做法是把画面拆解成 PowerPoint 原生对象：标题和正文写成文本框，卡片、色块和几何图形写成形状，流程线写成独立连接线，图标和复杂插图作为独立图片嵌回。交付的 PPTX 打开后可直接修改文案、调整配色、移动元素。

## 适用场景与处理策略

根据素材特点和目标选择策略，也可在同一份 PPT 中组合：

| 策略 | 适用素材 | 处理方式 |
| --- | --- | --- |
| **还原版式** | 页面结构完整、投屏清晰、对象边界明确的截图：教学课件、宣传海报、产品介绍、数据图表、流程图、已有幻灯片 | 按原图重建版式与对象，保留原文、专有名词、数字和流程方向；清晰且结构独立时一图一页 |
| **重排内容** | 手机长图、多屏聊天、网页、文章、密集表格 | 提取原图内容并按主题拆页；保留关键事实，模糊文字和看不清的数值标为待核对 |
| **保留关键画面** | 照片、产品图、操作界面等需要保留像素细节的素材 | 裁切关键画面作为独立图片放入 PPT；标题、说明、标注、框选和箭头用可编辑对象叠加 |

用户只说「转 PPT」时：页面结构完整且投屏清晰的默认还原，长图或密集内容默认重排并简要说明做法。指定了画幅、页数、顺序或还原要求时按指定交付；原图比例不同时等比缩放，并检查最密的一页是否可读。

## 工作流程

1. **记录**：页序、画幅、文案、边界、配色、叠放顺序及连接关系，低清内容标记待核对。
2. **重建**：标题、正文、标签用文本框；卡片、色块、几何图形用原生形状；流程线用独立连接线或箭头形状；对照表用原生表格或独立形状与文字；图标和复杂插图用独立图片，需要改色时用原生形状重建。
3. **核对**：重新打开生成的 PPTX 并逐页渲染，对照原图检查页序、中文、对齐、文字溢出、箭头方向、表格可读性和图片边缘；对承诺可编辑的对象在副本上实际验证改文字、改填充、移动箭头、独立选择图标。
4. **交付**：可编辑 PPTX，说明采用的策略、拆页对应关系、未辨认文字及以图片保留的内容。

## 安装与依赖

需要 Python 3.9+：

```bash
pip install python-pptx Pillow
```

获取脚本：

```bash
git clone https://github.com/justin-yu-space/wg-screenshot-to-ppt.git
cd wg-screenshot-to-ppt
```

首选构建器 `build.py`，基于 python-pptx。仓库另附等价的 Node 版 `build.mjs`，依赖 `@oai/artifact-tool`；该包是 OpenAI 内部依赖，通常无法从公共 npm 源获取，仅在环境已具备时使用。

## 清单格式

清单是描述每页画布与对象的 JSON。坐标单位统一为 **EMU**：1 英寸 = 914400 EMU；按 96 DPI 换算，1 像素 = 9525 EMU；1920×1080 像素 = 18288000×12192000 EMU。字号 `size` 单位为磅。

根节点包含 `width`、`height`、可选 `font` 和按原图顺序排列的 `slides`。每页写 `background` 和按叠放顺序排列的 `objects`。对象通用字段为 `name`、`x`、`y`、`w`、`h`，按类型补充：

| 类型 | 字段 |
| --- | --- |
| `text` | `text`、`size`（磅）、`color`，可选 `bold`、`wrap`（`square` 默认自动换行 / `none` 不换行）、`align`（`left`/`center`/`right`）、`valign`（`top`/`middle` 默认/`bottom`） |
| `shape` | `fill`，可选 `stroke`（六位 HEX）、`radius`（EMU 圆角半径） |
| `arrow` | `color`，可选 `head`（`arrow` 默认 / `triangle` / `stealth` / `diamond` / `oval` / `none`）；终点为 `(x+w, y+h)`，可用负宽高表达反向 |
| `image` | `file`：相对清单目录的图片路径，支持 PNG/JPG/JPEG/GIF/BMP/WEBP |

颜色为六位 HEX（可带 `#`），`none` 或 `null` 表示无填充。`image.file` 只能指向清单所在目录内的文件。

示例（16:9 画布，1280×720 像素）：

```json
{
  "width": 12192000,
  "height": 6858000,
  "font": "Microsoft YaHei",
  "slides": [
    {
      "background": "#FFFFFF",
      "objects": [
        { "type": "text", "name": "标题", "x": 952500, "y": 457200,
          "w": 10287000, "h": 685800, "text": "示例标题", "size": 28,
          "color": "#1F2937", "bold": true, "wrap": "square", "align": "left" },
        { "type": "shape", "name": "卡片", "x": 952500, "y": 1371600,
          "w": 3429000, "h": 2286000, "fill": "#E5E7EB", "stroke": "#9CA3AF", "radius": 152400 },
        { "type": "arrow", "name": "流程线", "x": 4572000, "y": 2514600,
          "w": 1828800, "h": 0, "color": "#2563EB", "head": "arrow" },
        { "type": "image", "name": "图标", "x": 6858000, "y": 1371600,
          "w": 1143000, "h": 1143000, "file": "icon.png" }
      ]
    }
  ]
}
```

可运行的完整示例见 [`examples/manifest.json`](examples/manifest.json)。

## 构建与裁切

```bash
# 生成可编辑 PPTX
python scripts/build.py 清单.json 成品.pptx

# Node 版（需 @oai/artifact-tool）
node scripts/build.mjs 清单.json 成品.pptx

# 裁切图标：--box 为 x y w h 裁切区域，--size 输出方形画布边长
python scripts/crop_icon.py source.png icon.png --box 10 20 64 64 --background FFFFFF --foreground 2563EB --size 128
```

`crop_icon.py` 对单色图标可按背景色与前景色恢复边缘透明度（`--background` 与 `--foreground` 须同时给出，六位 HEX）。

两个构建器内置清单校验：字段缺失、类型错误、数值非法、图片路径逃逸清单目录或文件不存在，都会报出可定位到 `slides[i].objects[j]` 的错误。

## 文件说明

| 路径 | 用途 |
| --- | --- |
| `SKILL.md` | 技能定义：策略选择、重建与核对流程、构建工具说明 |
| `scripts/build.py` | 首选构建器（python-pptx） |
| `scripts/build.mjs` | 备选构建器（@oai/artifact-tool） |
| `scripts/crop_icon.py` | 图标裁切与透明背景恢复 |
| `examples/` | 可运行的示例清单、素材与输出成品 |
