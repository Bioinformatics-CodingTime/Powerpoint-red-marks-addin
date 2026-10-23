const fs = require("node:fs");
const path = require("node:path");

const root = __dirname;
const output = Buffer.concat([
  Buffer.from('window.MathJax = {svg: {fontCache: "none"}, startup: {typeset: false}};\n'),
  fs.readFileSync(path.join(root, "node_modules/mathjax/es5/tex-svg.js")),
  Buffer.from("\n"),
  fs.readFileSync(path.join(root, "src/commands.core.js")),
]);
const target = path.join(root, "src/commands.js");
if (process.argv[2] === "--check") {
  if (!fs.existsSync(target) || !fs.readFileSync(target).equals(output)) {
    throw new Error("commands.js 与源码不一致；请运行 node build.js");
  }
} else {
  fs.writeFileSync(target, output);
}
