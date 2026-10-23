/* global Office, PowerPoint, MathJax */

Office.onReady(() => {});

async function insertRedArrow(event) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const arrow = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rightArrow,
        { left: 360, top: 245, width: 240, height: 50 }
      );

      arrow.fill.setSolidColor("#FF0000");
      arrow.lineFormat.color = "#FF0000";
      arrow.lineFormat.weight = 1.5;
      arrow.load("id");
      await context.sync();
      slide.setSelectedShapes([arrow.id]);
      await context.sync();
    });
  } catch (error) {
    console.error("插入红色箭头失败", error);
  } finally {
    event.completed();
  }
}

async function insertRedBox(event) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const box = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left: 300, top: 170, width: 360, height: 200 }
      );
      box.fill.clear();
      box.lineFormat.visible = true;
      box.lineFormat.color = "#FF0000";
      box.lineFormat.weight = 3;
      box.load("id");
      await context.sync();
      slide.setSelectedShapes([box.id]);
      await context.sync();
    });
  } catch (error) {
    console.error("插入红色方框失败", error);
  } finally {
    event.completed();
  }
}

async function insertWhiteBox(event) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const box = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left: 300, top: 170, width: 360, height: 200 }
      );
      box.fill.setSolidColor("#FFFFFF");
      box.lineFormat.visible = false;
      box.load("id");
      await context.sync();
      slide.setSelectedShapes([box.id]);
      await context.sync();
    });
  } catch (error) {
    console.error("插入纯白方框失败", error);
  } finally {
    event.completed();
  }
}

function toMarkdown(range) {
  const text = range.text.trim();
  if (!text) return "";
  const size = range.font.size;
  const heading = size >= 28 ? "# " : size >= 23 ? "## " : "";
  const prefix = range.paragraphFormat.bulletFormat.visible ? "- " : heading;
  const marker = range.font.bold ? "**" : range.font.italic ? "*" : "";
  return text.split(/\r?\n/).map((line) => line ? `${prefix}${marker}${line}${marker}` : "").join("\n");
}

async function copyNativeFormula(context, text) {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
    throw new Error("当前 PowerPoint 版本不支持导出原生公式。");
  }
  const exported = context.presentation.getSelectedSlides().exportAsBase64Presentation();
  await context.sync();
  return { formulaText: text, presentation: exported.value };
}

async function originalFormula(context, shape, text) {
  const tag = shape.tags.getItemOrNullObject("MARKDOWN_FORMULA");
  tag.load("value");
  await context.sync();
  if (!tag.isNullObject) {
    try {
      const saved = JSON.parse(tag.value);
      if (saved.raw === encodeURIComponent(text)) return saved.source;
    } catch { /* An edited or old formula uses the normal converter. */ }
  }
  return copyNativeFormula(context, text);
}

// shortcut: mixed formatting, tables, equations, and images stay plain text; upgrade for full Writage parity.
function fromMarkdown(markdown) {
  const scripts = {
    sub: { 0: "₀", 1: "₁", 2: "₂", 3: "₃", 4: "₄", 5: "₅", 6: "₆", 7: "₇", 8: "₈", 9: "₉", "+": "₊", "-": "₋", "=": "₌", "(": "₍", ")": "₎" },
    sup: { 0: "⁰", 1: "¹", 2: "²", 3: "³", 4: "⁴", 5: "⁵", 6: "⁶", 7: "⁷", 8: "⁸", 9: "⁹", "+": "⁺", "-": "⁻", "=": "⁼", "(": "⁽", ")": "⁾" },
  };
  const renderMath = (line) => line.replace(/\$([^$\n]+)\$/g, (source, expression) => {
    const converted = expression.replace(/([_^])(?:\{([^{}]+)\}|([0-9]))/g, (token, kind, grouped, single) => {
      const alphabet = kind === "_" ? scripts.sub : scripts.sup;
      const characters = [...(grouped || single)].map((character) => alphabet[character]);
      return characters.every(Boolean) ? characters.join("") : token;
    });
    return /[\\{}_\^]/.test(converted) ? source : converted;
  });
  const styles = [];
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  let text = "";
  for (const source of lines) {
    const heading = /^(#{1,6})\s+/.exec(source);
    const bullet = !heading && /^[-*+]\s+/.test(source);
    const line = renderMath(source.replace(/^(#{1,6}|[-*+])\s+/, ""));
    const start = text.length;
    let position = 0;
    const pattern = /(\*\*([^*]+)\*\*|\*([^*]+)\*)/g;
    let match;
    while ((match = pattern.exec(line))) {
      text += line.slice(position, match.index);
      const content = match[2] || match[3];
      styles.push({ start: text.length, length: content.length, bold: !!match[2], italic: !!match[3] });
      text += content;
      position = pattern.lastIndex;
    }
    text += line.slice(position) + "\n";
    if (heading || bullet) styles.push({ start, length: text.length - start - 1, heading: heading?.[1].length, bullet });
  }
  return { text: text.replace(/\n$/, ""), styles };
}

async function readClipboard() {
  let corrupted = false;
  try {
    const response = await fetch("/clipboard/read", { method: "POST" });
    if (response.ok) {
      const text = await response.text();
      if (!text.includes("\uFFFD")) return text;
      corrupted = true;
    }
  } catch (error) { console.warn("本机剪贴板接口不可用", error); }
  if (navigator.clipboard?.readText) {
    const text = await navigator.clipboard.readText();
    if (!text.includes("\uFFFD")) return text;
  }
  if (corrupted) throw new Error("剪贴板读取出现乱码；请重启本地加载项服务。");
  throw new Error("无法读取剪贴板；请重新启动本地加载项服务。");
}

async function writeClipboard(value) {
  try {
    const response = await fetch("/clipboard/write", {
      method: "POST", headers: { "Content-Type": "text/plain" }, body: value,
    });
    if (response.ok) return;
  } catch (error) { console.warn("本机剪贴板接口不可用", error); }
  if (navigator.clipboard?.writeText) return navigator.clipboard.writeText(value);
  throw new Error("无法写入剪贴板；请重新启动本地加载项服务。");
}

async function copyMarkdown(event) {
  try {
    const markdown = await PowerPoint.run(async (context) => {
      const selectedText = context.presentation.getSelectedTextRangeOrNullObject();
      selectedText.load("text");
      await context.sync();
      if (!selectedText.isNullObject && selectedText.text.trim()) {
        if (selectedText.text.includes("\uFDD0")) return originalFormula(context, selectedText.getParentTextFrame().getParentShape(), selectedText.text);
        selectedText.load("font/size,font/bold,font/italic,paragraphFormat/bulletFormat/visible");
        await context.sync();
        return toMarkdown(selectedText);
      }
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();
      const frames = shapes.items.map((shape) => shape.textFrame);
      frames.forEach((frame) => frame.load("hasText"));
      await context.sync();
      const textShapes = shapes.items.filter((shape, index) => frames[index].hasText);
      const ranges = textShapes.map((shape) => shape.textFrame.textRange);
      ranges.forEach((range) => range.load("text,font/size,font/bold,font/italic,paragraphFormat/bulletFormat/visible"));
      await context.sync();
      if (ranges.some((range) => range.text.includes("\uFDD0"))) {
        if (ranges.length !== 1) throw new Error("请一次只选中一个公式文本框。");
        return originalFormula(context, textShapes[0], ranges[0].text);
      }
      return ranges.map(toMarkdown).filter(Boolean).join("\n\n");
    });
    if (!markdown) throw new Error("请先选中幻灯片中的文字或文本框。");
    if (typeof markdown === "string") await writeClipboard(markdown);
    else {
      const response = await fetch("/formula/copy", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(markdown),
      });
      if (!response.ok) throw new Error(`公式转换失败：${await response.text()}`);
      await writeClipboard(await response.text());
    }
  } catch (error) {
    console.error("Copy Markdown 失败", error);
    if (typeof alert === "function") alert(`Copy Markdown 失败：${error.message}`);
  } finally {
    event.completed();
  }
}

async function insertFormula(expression, source) {
  const response = await fetch("/formula/render", {
    method: "POST", headers: { "Content-Type": "text/plain" }, body: expression,
  });
  if (!response.ok) throw new Error(`原生公式生成失败：${await response.text()}`);
  const presentation = await response.text();
  let targetId;
  let sourceId;
  let targetIndex;
  let sourceIndex;
  let oldShapeIds;
  await PowerPoint.run(async (context) => {
    const target = context.presentation.getSelectedSlides().getItemAt(0);
    const slides = context.presentation.slides;
    target.load("id");
    slides.load("items/id");
    target.shapes.load("items/id");
    await context.sync();
    targetId = target.id;
    targetIndex = slides.items.findIndex((slide) => slide.id === targetId) + 1;
    if (!targetIndex) throw new Error("无法定位当前幻灯片。");
    oldShapeIds = new Set(target.shapes.items.map((shape) => shape.id));
    const oldSlideIds = new Set(slides.items.map((slide) => slide.id));
    context.presentation.insertSlidesFromBase64(presentation, { targetSlideId: `${targetId}#`, formatting: "UseDestinationTheme" });
    context.presentation.setSelectedSlides([targetId]);
    await context.sync();
    slides.load("items/id");
    await context.sync();
    const added = slides.items.filter((slide) => !oldSlideIds.has(slide.id));
    if (added.length !== 1) throw new Error("临时公式页数量异常。");
    sourceId = added[0].id;
    sourceIndex = slides.items.findIndex((slide) => slide.id === sourceId) + 1;
  });
  try {
    const placed = await fetch("/formula/place", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ targetIndex, sourceIndex }),
    });
    if (!placed.ok) throw new Error(`原生公式插入失败：${await placed.text()}`);
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItem(targetId);
      const shapes = slide.shapes;
      shapes.load("items/id");
      await context.sync();
      const added = shapes.items.filter((shape) => !oldShapeIds.has(shape.id));
      if (added.length !== 1) throw new Error("无法唯一定位刚插入的公式。");
      added[0].textFrame.textRange.load("text");
      await context.sync();
      const raw = encodeURIComponent(added[0].textFrame.textRange.text);
      added[0].tags.add("MARKDOWN_FORMULA", JSON.stringify({ source, raw }));
      await context.sync();
      const tag = added[0].tags.getItem("MARKDOWN_FORMULA");
      tag.load("value");
      await context.sync();
      if (tag.value !== JSON.stringify({ source, raw })) throw new Error("公式原文没有保存成功。");
      slide.setSelectedShapes([added[0].id]);
      await context.sync();
    });
  } finally {
    if (sourceId) await PowerPoint.run(async (context) => {
      const temporary = context.presentation.slides.getItemOrNullObject(sourceId);
      await context.sync();
      if (!temporary.isNullObject) temporary.delete();
      await context.sync();
    });
  }
}

async function pasteMarkdown(event) {
  try {
    const markdown = await readClipboard();
    if (!markdown.trim()) throw new Error("剪贴板里没有文字。");
    if (markdown.length > 10000 || markdown.split(/\r?\n/).length > 100) {
      throw new Error("内容过长；请分段粘贴到多张幻灯片。");
    }
    const trimmed = markdown.trim();
    const formula = /^\$\$([\s\S]+?)\$\$$/.exec(trimmed) || /^\$([^$\n]+)\$$/.exec(trimmed);
    const bareFormula = !/[\r\n]/.test(trimmed) && (/\\[A-Za-z]+/.test(trimmed) || /[_^](?:\{[^{}]+\}|[A-Za-z0-9])/.test(trimmed));
    if (formula || bareFormula) {
      await insertFormula(formula ? formula[1] : trimmed, trimmed);
      return;
    }
    const rendered = fromMarkdown(markdown);
    if (rendered.styles.length > 500) throw new Error("格式片段过多；请分段粘贴。");
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shape = slide.shapes.addTextBox(rendered.text, { left: 80, top: 80, width: 800, height: 400 });
      shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeTextToFitShape;
      shape.textFrame.textRange.font.size = 18;
      for (const style of rendered.styles) {
        if (!style.length) continue;
        const range = shape.textFrame.textRange.getSubstring(style.start, style.length);
        if (style.heading) {
          range.font.bold = true;
          range.font.size = Math.max(18, 36 - style.heading * 4);
        }
        if (style.bullet) range.paragraphFormat.bulletFormat.visible = true;
        if (style.bold) range.font.bold = true;
        if (style.italic) range.font.italic = true;
      }
      shape.load("id");
      await context.sync();
      slide.setSelectedShapes([shape.id]);
      await context.sync();
    });
  } catch (error) {
    console.error("Paste Markdown 失败", error);
    if (typeof alert === "function") alert(`Paste Markdown 失败：${error.message}`);
  } finally {
    event.completed();
  }
}

Office.actions.associate("insertRedArrow", insertRedArrow);
Office.actions.associate("insertRedBox", insertRedBox);
Office.actions.associate("insertWhiteBox", insertWhiteBox);
Office.actions.associate("copyMarkdown", copyMarkdown);
Office.actions.associate("pasteMarkdown", pasteMarkdown);
