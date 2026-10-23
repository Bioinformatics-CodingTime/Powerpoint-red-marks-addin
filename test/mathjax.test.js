const assert = require("node:assert/strict");

require("mathjax").init({
  loader: { load: ["input/tex", "output/svg"] },
  svg: { fontCache: "none" },
  startup: { typeset: false },
}).then(async (MathJax) => {
  const tex = String.raw`P_{\mathrm{emp}}=\frac{n_{\mathrm{null}\geq\mathrm{obs}}+1}{N_{\mathrm{null}}+1}`;
  const formula = await MathJax.tex2svgPromise(tex, { display: true });
  const svg = MathJax.startup.adaptor.serializeXML(formula);
  assert.match(svg, /<svg\b/);
  assert.match(svg, /data-mml-node="mfrac"/);
  assert.doesNotMatch(svg, /data-mml-node="merror"/);
  console.log("MathJax 独立解析检查通过（非 PowerPoint 插入验证）。");
}).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
