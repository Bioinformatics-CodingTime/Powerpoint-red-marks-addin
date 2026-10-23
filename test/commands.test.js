const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const actions = {};
const created = [];
const selected = [];
const formatted = [];
const formulas = [];
const insertedFormulaExpressions = [];
const alerts = [];
let clipboard = "";
let formulaShape;
const rawFormula = "\uFDD0𝑃\uFDEEemp\uFDEF=\uFDD0\uFDD0𝑛\uFDEEnull≥obs\uFDEF+1\uFDEE\uFDD0𝑁\uFDEEnull\uFDEF+1\uFDEF";
function tags() {
  const values = new Map();
  return {
    add(key, value) { values.set(key, value); },
    getItem(key) { return { value: values.get(key), load() {} }; },
    getItemOrNullObject(key) { return { value: values.get(key), isNullObject: !values.has(key), load() {} }; },
  };
}
const selectedText = {
  text: "Result",
  isNullObject: false,
  font: { size: 30, bold: true, italic: false },
  paragraphFormat: { bulletFormat: { visible: false } },
  load() {},
  getParentTextFrame: () => ({ getParentShape: () => formulaShape || { tags: tags() } }),
};
const slide = {
  id: "opaque-target-id",
  load() {},
  shapes: {
    items: created,
    load() {},
    addGeometricShape(type, options) {
      const shape = {
        id: `shape-${created.length + 1}`,
        type,
        options,
        fill: {
          clear() { shape.fill.cleared = true; },
          setSolidColor(color) { shape.fill.color = color; },
        },
        lineFormat: {},
        load() {},
      };
      created.push(shape);
      return shape;
    },
    addTextBox(text, options) {
      const textRange = {
        font: {},
        getSubstring(start, length) {
          const range = { start, length, font: {}, paragraphFormat: { bulletFormat: {} } };
          formatted.push(range);
          return range;
        },
      };
      const shape = {
        id: `shape-${created.length + 1}`,
        type: "TextBox", text, options,
        textFrame: { textRange },
        load() {},
      };
      created.push(shape);
      return shape;
    },
  },
  setSelectedShapes(ids) { selected.push([...ids]); },
};
const context = {
  presentation: {
    getSelectedSlides: () => ({ getItemAt: () => slide, exportAsBase64Presentation: () => ({ value: "test-pptx-base64" }) }),
    slides: {
      items: [slide],
      load() {},
      getItem(id) { assert.equal(id, "opaque-target-id"); return slide; },
      getItemOrNullObject(id) {
        return { isNullObject: !this.items.some((item) => item.id === id), load() {}, delete: () => { this.items = this.items.filter((item) => item.id !== id); } };
      },
    },
    insertSlidesFromBase64(value, options) {
      assert.equal(value, "test-rendered-pptx");
      assert.equal(options.targetSlideId, "opaque-target-id#");
      assert.equal(options.formatting, "UseDestinationTheme");
      this.slides.items.push({ id: "opaque-source-id" });
    },
    setSelectedSlides(ids) { assert.deepEqual([...ids], ["opaque-target-id"]); },
    getSelectedTextRangeOrNullObject: () => selectedText,
    getSelectedShapes: () => ({
      items: [{ textFrame: {
        hasText: true,
        textRange: {
          text: "item",
          font: { size: 18, bold: false, italic: false },
          paragraphFormat: { bulletFormat: { visible: true } },
          load() {},
        },
        load() {},
      } }],
      load() {},
    }),
  },
  async sync() {},
};
const sandbox = {
  console,
  alert(message) { alerts.push(message); },
  Office: {
    onReady() {},
    actions: { associate(name, fn) { actions[name] = fn; } },
    context: {
      requirements: { isSetSupported: () => true },
    },
  },
  MathJax: {
    startup: { promise: Promise.resolve() },
    async tex2svgPromise(tex, options) {
      formulas.push({ tex, display: options.display });
      const attributes = { viewBox: "0 -500 10000 3000" };
      const svg = {
        getAttribute(name) { return attributes[name]; },
        querySelector() { return null; },
        setAttribute(name, value) { attributes[name] = value; },
        get outerHTML() { return `<svg xmlns="${attributes.xmlns}" width="${attributes.width}" height="${attributes.height}" viewBox="${attributes.viewBox}"><path fill="currentColor"/></svg>`; },
      };
      return { querySelector: () => svg };
    },
  },
  Image: class {
    set src(value) {
      assert.match(value, /^data:image\/svg\+xml;base64,/);
      this.onload();
    }
  },
  document: {
    createElement(name) {
      assert.equal(name, "canvas");
      return {
        getContext(kind) {
          assert.equal(kind, "2d");
          return { drawImage() {} };
        },
        toDataURL(kind) {
          assert.equal(kind, "image/png");
          return `data:image/png;base64,${btoa("PNG")}`;
        },
      };
    },
  },
  TextEncoder,
  btoa,
  PowerPoint: {
    GeometricShapeType: { rightArrow: "RightArrow", rectangle: "Rectangle" },
    ShapeAutoSize: { autoSizeTextToFitShape: "AutoSizeTextToFitShape" },
    run: (fn) => fn(context),
  },
  navigator: {},
  fetch: async (url, options) => {
    if (url === "/clipboard/write" && options.method === "POST") {
      clipboard = options.body;
      return { ok: true };
    }
    if (url === "/clipboard/read" && options.method === "POST") {
      return { ok: true, text: async () => clipboard };
    }
    if (url === "/formula/render" && options.method === "POST") {
      insertedFormulaExpressions.push(options.body);
      return { ok: true, text: async () => "test-rendered-pptx" };
    }
    if (url === "/formula/place" && options.method === "POST") {
      assert.deepEqual(JSON.parse(options.body), { targetIndex: 1, sourceIndex: 2 });
      formulaShape = { id: `formula-${created.length + 1}`, tags: tags(), textFrame: { textRange: { text: rawFormula, load() {} } } };
      created.push(formulaShape);
      return { ok: true, text: async () => "AppleScript linear text differs from Office.js" };
    }
    if (url === "/formula/copy" && options.method === "POST") {
      const request = JSON.parse(options.body);
      assert.equal(request.presentation, "test-pptx-base64");
      assert.equal(request.formulaText, selectedText.text);
      return { ok: true, text: async () => "$$P_{\\text{emp}} = \\frac{n_{\\text{null} \\geq \\text{obs}} + 1}{N_{\\text{null}} + 1}$$" };
    }
    throw new Error("Unexpected clipboard call");
  },
};

vm.runInNewContext(
  fs.readFileSync(path.join(__dirname, "../src/commands.core.js"), "utf8"),
  sandbox
);

(async () => {
  let completed = 0;
  await actions.insertRedArrow({ completed() { completed += 1; } });
  await actions.insertRedBox({ completed() { completed += 1; } });
  await actions.insertWhiteBox({ completed() { completed += 1; } });

  assert.equal(created[0].type, "RightArrow");
  assert.equal(created[0].fill.color, "#FF0000");
  assert.equal(created[1].type, "Rectangle");
  assert.equal(created[1].fill.cleared, true);
  assert.equal(created[1].lineFormat.color, "#FF0000");
  assert.equal(created[1].lineFormat.weight, 3);
  assert.equal(created[2].type, "Rectangle");
  assert.equal(created[2].fill.color, "#FFFFFF");
  assert.equal(created[2].lineFormat.visible, false);
  await actions.copyMarkdown({ completed() { completed += 1; } });
  assert.equal(clipboard, "# **Result**");
  selectedText.text = rawFormula;
  await actions.copyMarkdown({ completed() { completed += 1; } });
  assert.equal(clipboard, "$$P_{\\text{emp}} = \\frac{n_{\\text{null} \\geq \\text{obs}} + 1}{N_{\\text{null}} + 1}$$");
  selectedText.isNullObject = true;
  selectedText.text = "";
  await actions.copyMarkdown({ completed() { completed += 1; } });
  assert.equal(clipboard, "- item");
  clipboard = "# Title\n- item with **bold** and *italic*";
  await actions.pasteMarkdown({ completed() { completed += 1; } });
  assert.equal(created[3].text, "Title\nitem with bold and italic");
  assert.equal(created[3].textFrame.autoSizeSetting, "AutoSizeTextToFitShape");
  assert.equal(formatted.find((range) => range.start === 0).font.size, 32);
  assert.equal(formatted.some((range) => range.paragraphFormat.bulletFormat.visible), true);
  assert.equal(formatted.find((range) => range.font.italic).length, 6);
  clipboard = "$p_1=p_2=10^{-4}$";
  await actions.pasteMarkdown({ completed() { completed += 1; } });
  clipboard = "$$P_{\\mathrm{emp}}=\\frac{n_{\\mathrm{null}\\geq\\mathrm{obs}}+1}{N_{\\mathrm{null}}+1}$$";
  await actions.pasteMarkdown({ completed() { completed += 1; } });
  clipboard = "P_{\\mathrm{emp}}=\\frac{n_{\\mathrm{null}\\geq\\mathrm{obs}}+1}{N_{\\mathrm{null}}+1}";
  await actions.pasteMarkdown({ completed() { completed += 1; } });
  clipboard = "\uFFFDbad";
  sandbox.navigator.clipboard = { readText: async () => "$q=1$" };
  await actions.pasteMarkdown({ completed() { completed += 1; } });
  sandbox.navigator.clipboard = undefined;
  await assert.rejects(sandbox.readClipboard(), /剪贴板读取出现乱码/);
  assert.equal(alerts.length, 0);
  assert.deepEqual(insertedFormulaExpressions, [
    "p_1=p_2=10^{-4}",
    "P_{\\mathrm{emp}}=\\frac{n_{\\mathrm{null}\\geq\\mathrm{obs}}+1}{N_{\\mathrm{null}}+1}",
    "P_{\\mathrm{emp}}=\\frac{n_{\\mathrm{null}\\geq\\mathrm{obs}}+1}{N_{\\mathrm{null}}+1}",
    "q=1",
  ]);
  assert.equal(context.presentation.slides.items.length, 1);
  selectedText.isNullObject = false;
  selectedText.text = rawFormula;
  await actions.copyMarkdown({ completed() { completed += 1; } });
  assert.equal(clipboard, "$q=1$");
  selectedText.text += "2";
  await actions.copyMarkdown({ completed() { completed += 1; } });
  assert.equal(clipboard, "$$P_{\\text{emp}} = \\frac{n_{\\text{null} \\geq \\text{obs}} + 1}{N_{\\text{null}} + 1}$$");
  assert.equal(formulas.length, 0);
  assert.equal(created.length, 8);
  assert.equal(sandbox.fromMarkdown("$\\frac{1}{2}$").text, "$\\frac{1}{2}$");
  assert.deepEqual(selected, [["shape-1"], ["shape-2"], ["shape-3"], ["shape-4"], ["formula-5"], ["formula-6"], ["formula-7"], ["formula-8"]]);
  assert.equal(completed, 13);
  console.log("按钮逻辑模拟检查通过。");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
