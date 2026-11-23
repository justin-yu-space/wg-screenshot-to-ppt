import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const [manifestPath, outputPath] = process.argv.slice(2);
if (!manifestPath || !outputPath) throw new Error("Usage: node build.mjs manifest.json output.pptx");
const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
const sourceDir = path.dirname(path.resolve(manifestPath));

const OBJECT_TYPES = new Set(["text", "shape", "arrow", "image"]);
const MIME = {
  ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
  ".gif": "image/gif", ".bmp": "image/bmp", ".webp": "image/webp",
};

function fail(message) { throw new Error(`清单错误: ${message}`); }
function number(obj, key, where, { allowNegative = false, positive = false } = {}) {
  if (!(key in obj)) fail(`${where} 缺少必填字段 ${key}`);
  const value = obj[key];
  if (typeof value !== "number" || !Number.isFinite(value)) fail(`${where}.${key} 必须是数值`);
  if (positive && value <= 0) fail(`${where}.${key} 必须为正数`);
  if (!allowNegative && value < 0) fail(`${where}.${key} 不能为负数`);
  return value;
}
function color(value, where) {
  if (value === null || value === undefined) return undefined;
  if (typeof value !== "string") fail(`${where} 必须是颜色字符串或 null`);
  if (/^(none|transparent)$/i.test(value.trim())) return undefined;
  const s = value.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(s)) fail(`${where} 必须是六位 HEX（可带 #），实际为 ${value}`);
  return `#${s}`;
}

if (!manifest || typeof manifest !== "object" || Array.isArray(manifest)) fail("根节点必须是对象");
number(manifest, "width", "清单", { positive: true });
number(manifest, "height", "清单", { positive: true });
if (manifest.font !== undefined && typeof manifest.font !== "string") fail("清单.font 必须是字符串");
if (!Array.isArray(manifest.slides) || !manifest.slides.length) fail("清单.slides 必须是非空数组");

for (const [si, page] of manifest.slides.entries()) {
  const where = `slides[${si}]`;
  if (!page || typeof page !== "object") fail(`${where} 必须是对象`);
  if (!("background" in page)) fail(`${where} 缺少必填字段 background`);
  color(page.background, `${where}.background`);
  if (!Array.isArray(page.objects)) fail(`${where}.objects 必须是数组`);
  for (const [oi, item] of page.objects.entries()) {
    const w = `${where}.objects[${oi}]`;
    if (!item || typeof item !== "object") fail(`${w} 必须是对象`);
    if (typeof item.name !== "string") fail(`${w}.name 必须是字符串`);
    if (!OBJECT_TYPES.has(item.type)) fail(`${w}.type 不是支持的类型: ${item.type}`);
    const negOK = item.type === "arrow";
    for (const key of ["x", "y", "w", "h"]) number(item, key, w, { allowNegative: negOK || key === "x" || key === "y" });
    if (item.type === "text") {
      if (typeof item.text !== "string") fail(`${w}.text 必须是字符串`);
      number(item, "size", w, { positive: true });
      color(item.color, `${w}.color`);
      if (item.wrap !== undefined && !["square", "none"].includes(item.wrap)) fail(`${w}.wrap 必须是 'square' 或 'none'`);
      if (item.align !== undefined && !["left", "center", "right"].includes(item.align)) fail(`${w}.align 不是合法值`);
    } else if (item.type === "shape") {
      color(item.fill, `${w}.fill`);
      if (item.stroke !== undefined) color(item.stroke, `${w}.stroke`);
      if (item.radius !== undefined) number(item, "radius", w);
    } else if (item.type === "arrow") {
      color(item.color, `${w}.color`);
      if (item.head !== undefined && !["arrow", "triangle", "stealth", "diamond", "oval", "none"].includes(item.head)) fail(`${w}.head 不是合法值`);
    } else if (item.type === "image") {
      if (typeof item.file !== "string" || !item.file) fail(`${w}.file 必须是非空字符串`);
      const file = path.resolve(sourceDir, item.file);
      if (!file.startsWith(sourceDir + path.sep)) fail(`${w}.file 逃逸清单目录`);
      await fs.access(file);
    }
  }
}

const deck = Presentation.create({ slideSize: { width: manifest.width, height: manifest.height } });

for (const page of manifest.slides) {
  const slide = deck.slides.add();
  slide.background.fill = page.background ?? "none";
  for (const item of page.objects) {
    const position = { left: item.x, top: item.y, width: item.w, height: item.h };
    if (item.type === "image") {
      const file = path.resolve(sourceDir, item.file);
      const contentType = MIME[path.extname(file).toLowerCase()] ?? "image/png";
      const blob = await fs.readFile(file);
      slide.images.add({ blob: new Uint8Array(blob), contentType, alt: item.name, position });
    } else if (item.type === "arrow") {
      slide.shapes.add({ name: item.name, geometry: "connector", kind: "straight", position,
        line: { fill: item.color ?? "#000000", width: 3 },
        tail: { type: item.head === "none" ? "none" : (item.head ?? "arrow") } });
    } else if (item.type === "shape") {
      slide.shapes.add({ name: item.name, geometry: item.radius ? "roundRect" : "rect", position,
        fill: item.fill ?? "none", borderRadius: item.radius,
        line: item.stroke ? { fill: item.stroke, width: 2 } : undefined });
    } else if (item.type === "text") {
      const shape = slide.shapes.add({ name: item.name, geometry: "textbox", position, fill: "none" });
      shape.text = item.text;
      shape.text.style = { typeface: manifest.font ?? "Microsoft YaHei", fontSize: item.size,
        color: item.color ?? "#000000", bold: !!item.bold, insets: 0,
        verticalAlignment: item.valign ?? "middle", wrap: item.wrap ?? "square",
        align: item.align ?? "left" };
    }
  }
}
await (await PresentationFile.exportPptx(deck)).save(outputPath);
