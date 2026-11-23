#!/usr/bin/env python3
"""用 python-pptx 从清单 JSON 构建可编辑 PPTX。

用法: python build.py <manifest.json> <output.pptx>

清单格式与 scripts/build.mjs 保持一致，两个构建器接受同一份输入。
坐标单位 width/height/x/y/w/h 均为 EMU（与 @oai/artifact-tool 一致）。
文本对象额外支持可选 wrap 字段: "square"（默认，自动换行）或 "none"（不换行）。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

OBJECT_TYPES = {"text", "shape", "arrow", "image"}
ARROW_TYPES = {"arrow", "triangle", "stealth", "diamond", "oval", "none"}


class ManifestError(ValueError):
    """清单结构或字段错误，错误信息可直接定位到字段。"""


def _hex(value: str, field: str) -> tuple[int, int, int]:
    s = value.strip().lstrip("#")
    if len(s) != 6:
        raise ManifestError(f"{field} 必须是六位 HEX（可带 #），实际为 {value!r}")
    try:
        return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError as exc:
        raise ManifestError(f"{field} 颜色含非法字符: {value!r}") from exc


def parse_color(value, field: str):
    """返回 RGBColor 或 None（none/transparent/null 表示无填充）。"""
    if value is None:
        return None
    if not isinstance(value, str):
        raise ManifestError(f"{field} 必须是颜色字符串或 null")
    if value.strip().lower() in {"none", "transparent"}:
        return None
    r, g, b = _hex(value, field)
    return RGBColor(r, g, b)


def _number(obj: dict, key: str, where: str, *, allow_negative: bool = False) -> int:
    if key not in obj:
        raise ManifestError(f"{where} 缺少必填字段 {key!r}")
    value = obj[key]
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ManifestError(f"{where}.{key} 必须是数值，实际为 {type(value).__name__}")
    rounded = round(value)
    if not allow_negative and rounded <= 0:
        raise ManifestError(f"{where}.{key} 必须为正数，实际为 {value}")
    return rounded


def validate_manifest(data, source_dir: Path) -> None:
    """校验清单结构；任何错误抛出可定位到字段的 ManifestError。"""
    if not isinstance(data, dict):
        raise ManifestError("清单根节点必须是对象")
    _number(data, "width", "清单")
    _number(data, "height", "清单")
    if "font" in data and not isinstance(data["font"], str):
        raise ManifestError("清单.font 必须是字符串")
    slides = data.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ManifestError("清单.slides 必须是非空数组")

    for si, page in enumerate(slides):
        where = f"slides[{si}]"
        if not isinstance(page, dict):
            raise ManifestError(f"{where} 必须是对象")
        if "background" not in page:
            raise ManifestError(f"{where} 缺少必填字段 'background'")
        parse_color(page["background"], f"{where}.background")
        objects = page.get("objects")
        if not isinstance(objects, list):
            raise ManifestError(f"{where}.objects 必须是数组")
        for oi, item in enumerate(objects):
            _validate_object(item, f"{where}.objects[{oi}]", source_dir)


def _validate_object(item, where: str, source_dir: Path) -> None:
    if not isinstance(item, dict):
        raise ManifestError(f"{where} 必须是对象")
    if "name" not in item or not isinstance(item["name"], str):
        raise ManifestError(f"{where}.name 必须是字符串")
    kind = item.get("type")
    if kind not in OBJECT_TYPES:
        raise ManifestError(f"{where}.type 必须是 {sorted(OBJECT_TYPES)} 之一，实际为 {kind!r}")

    allow_negative = kind == "arrow"
    for key in ("x", "y", "w", "h"):
        _number(item, key, where, allow_negative=allow_negative or key in ("x", "y"))

    if kind == "text":
        if "text" not in item or not isinstance(item["text"], str):
            raise ManifestError(f"{where}.text 必须是字符串")
        _number(item, "size", where)
        parse_color(item.get("color"), f"{where}.color")
        for flag in ("bold",):
            if flag in item and not isinstance(item[flag], bool):
                raise ManifestError(f"{where}.{flag} 必须是布尔值")
        if "wrap" in item and item["wrap"] not in ("square", "none"):
            raise ManifestError(f"{where}.wrap 必须是 'square' 或 'none'")
        if "align" in item and item["align"] not in ("left", "center", "right"):
            raise ManifestError(f"{where}.align 必须是 'left'/'center'/'right'")
    elif kind == "shape":
        parse_color(item.get("fill"), f"{where}.fill")
        if "stroke" in item:
            parse_color(item["stroke"], f"{where}.stroke")
        if "radius" in item:
            _number(item, "radius", where)
    elif kind == "arrow":
        parse_color(item.get("color"), f"{where}.color")
        if "head" in item and item["head"] not in ARROW_TYPES:
            raise ManifestError(f"{where}.head 必须是 {sorted(ARROW_TYPES)} 之一")
    elif kind == "image":
        rel = item.get("file")
        if not isinstance(rel, str) or not rel:
            raise ManifestError(f"{where}.file 必须是非空字符串")
        target = (source_dir / rel).resolve()
        try:
            target.relative_to(source_dir)
        except ValueError as exc:
            raise ManifestError(f"{where}.file 逃逸清单目录: {rel!r}") from exc
        if not target.is_file():
            raise ManifestError(f"{where}.file 不存在: {target}")
        try:
            with Image.open(target) as img:
                img.verify()
        except Exception as exc:
            raise ManifestError(f"{where}.file 不是可识别图片: {target.name} ({exc})") from exc


def _fill(shape, color) -> None:
    if color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = color


def _set_tail_arrow(connector, arrow_type: str) -> None:
    """python-pptx 无箭头 API，直接写 DrawingML 的 a:tailEnd。"""
    if arrow_type == "none":
        return
    ln = connector.line._get_or_add_ln()
    tail = ln.find(qn("a:tailEnd"))
    if tail is None:
        tail = ln.makeelement(qn("a:tailEnd"), {})
        ln.append(tail)
    tail.set("type", "triangle" if arrow_type == "arrow" else arrow_type)
    tail.set("w", "med")
    tail.set("len", "med")


def _style_text(shape, item: dict, font: str) -> None:
    tf = shape.text_frame
    tf.text = item["text"]
    tf.word_wrap = item.get("wrap", "square") != "none"
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
    tf.vertical_anchor = anchor.get(item.get("valign", "middle"), MSO_ANCHOR.MIDDLE)
    alignment = {"left": 1, "center": 2, "right": 3}.get(item.get("align", "left"), 1)
    color = parse_color(item.get("color"), "text.color")
    for para in tf.paragraphs:
        para.alignment = alignment
        for run in para.runs:
            run.font.name = font
            run.font.size = Pt(item["size"])
            run.font.bold = bool(item.get("bold"))
            if color is not None:
                run.font.color.rgb = color


def build(manifest_path: Path, output_path: Path) -> None:
    source_dir = manifest_path.resolve().parent
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    validate_manifest(data, source_dir)

    deck = Presentation()
    deck.slide_width = Emu(data["width"])
    deck.slide_height = Emu(data["height"])
    font = data.get("font", "Microsoft YaHei")

    for page in data["slides"]:
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        bg = parse_color(page["background"], "background")
        if bg is None:
            slide.background.fill.background()
        else:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = bg

        for item in page["objects"]:
            pos = (Emu(item["x"]), Emu(item["y"]), Emu(item["w"]), Emu(item["h"]))
            kind = item["type"]
            if kind == "text":
                shape = slide.shapes.add_textbox(*pos)
                shape.name = item["name"]
                _style_text(shape, item, font)
            elif kind == "shape":
                geometry = MSO_SHAPE.ROUNDED_RECTANGLE if item.get("radius") else MSO_SHAPE.RECTANGLE
                shape = slide.shapes.add_shape(geometry, *pos)
                shape.name = item["name"]
                _fill(shape, parse_color(item.get("fill"), "shape.fill"))
                if item.get("radius"):
                    # python-pptx 圆角 adjustment 为短边比例，限制在 0-0.5
                    shape.adjustments[0] = min(0.5, max(0.0, item["radius"] / max(item["w"], item["h"])))
                if "stroke" in item:
                    stroke = parse_color(item["stroke"], "shape.stroke")
                    if stroke is None:
                        shape.line.fill.background()
                    else:
                        shape.line.color.rgb = stroke
                        shape.line.width = Pt(2)
                else:
                    shape.line.fill.background()
            elif kind == "arrow":
                connector = slide.shapes.add_connector(
                    MSO_CONNECTOR.STRAIGHT,
                    Emu(item["x"]), Emu(item["y"]),
                    Emu(item["x"] + item["w"]), Emu(item["y"] + item["h"]),
                )
                connector.name = item["name"]
                color = parse_color(item.get("color"), "arrow.color")
                if color is not None:
                    connector.line.color.rgb = color
                connector.line.width = Pt(3)
                _set_tail_arrow(connector, item.get("head", "arrow"))
            elif kind == "image":
                target = (source_dir / item["file"]).resolve()
                picture = slide.shapes.add_picture(str(target), *pos[:2], width=pos[2], height=pos[3])
                picture.name = item["name"]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    deck.save(str(output_path))


def main() -> None:
    parser = argparse.ArgumentParser(description="从清单 JSON 构建可编辑 PPTX")
    parser.add_argument("manifest", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    try:
        build(args.manifest, args.output)
    except (ManifestError, json.JSONDecodeError, OSError) as exc:
        print(f"构建失败: {exc}", file=sys.stderr)
        sys.exit(1)
    print(f"已生成: {args.output}")


if __name__ == "__main__":
    main()
