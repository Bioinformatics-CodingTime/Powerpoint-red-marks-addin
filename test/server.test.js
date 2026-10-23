const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

let handler;
const calls = [];
let script = "";
vm.runInNewContext(fs.readFileSync(path.join(__dirname, "../server.js"), "utf8"), {
  __dirname: path.join(__dirname, ".."),
  Buffer,
  URL,
  console,
  process: { env: { HOME: "/tmp" }, execPath: "/mock/node" },
  require(name) {
    if (name === "node:fs") return {
      readFileSync: (file) => file.endsWith("formula.pptx") ? Buffer.from("pptx") : "certificate",
      createReadStream: () => ({ pipe(response) { response.body = "script"; } }),
      mkdtempSync: () => "/private/tmp/ppt-markdown-test",
      writeFileSync: (_file, value) => { script = value; },
      rmSync() {},
    };
    if (name === "node:https") return { createServer(_options, callback) {
      handler = callback;
      return { listen() {} };
    } };
    if (name === "node:child_process") return { execFileSync(command, args, options) {
      calls.push([command, args, options]);
      if (command === "/usr/bin/unzip") return "<m:oMath";
      if (command === "/usr/bin/osascript") return "native formula\n";
      return args?.encoding === "utf8" ? "clipboard text" : Buffer.from("clipboard text");
    } };
    return require(name);
  },
});

function request(url, origin, contentType, body, method = "POST") {
  const input = new EventEmitter();
  input.url = url;
  input.method = method;
  input.headers = { origin, "content-type": contentType };
  input.destroy = () => {};
  const output = {
    writeHead(status, headers) { this.status = status; this.headers = headers; return this; },
    end(value) { this.body = value; return this; },
  };
  handler(input, output);
  if (body !== undefined) {
    input.emit("data", Buffer.from(body));
    input.emit("end");
  }
  return output;
}

assert.equal(request("/clipboard/read", "https://evil.example").status, 403);
assert.equal(request("/formula/copy", "https://evil.example", "application/json").status, 403);
assert.equal(request("/formula/render", "https://evil.example", "text/plain").status, 403);
assert.equal(request("/formula/place", "https://evil.example", "application/json").status, 403);
assert.equal(calls.length, 0);
assert.equal(request("/clipboard/read", "https://localhost:3000").body, "clipboard text");
assert.equal(calls.at(-1)[1].env.LANG, "en_US.UTF-8");
assert.equal(calls.at(-1)[1].env.LC_ALL, "en_US.UTF-8");
assert.equal(request("/clipboard/write", "https://localhost:3000", "application/json").status, 415);
assert.equal(request("/clipboard/write", "https://localhost:3000", "text/plain", "hello").status, 204);
assert.equal(calls.at(-1)[0], "/usr/bin/pbcopy");
assert.equal(calls.at(-1)[1].input.toString(), "hello");
assert.equal(calls.at(-1)[1].env.LC_ALL, "en_US.UTF-8");
assert.equal(request("/formula/render", "https://localhost:3000", "text/plain", "x^2").body, Buffer.from("pptx").toString("base64"));
assert.equal(request("/formula/place", "https://localhost:3000", "application/json", '{"targetIndex":1,"sourceIndex":2}').body, "native formula");
assert.match(script, /set sourceSlide to slide 2 of targetPresentation/);
assert.match(script, /FORMULA_TARGET_VIEW_MISMATCH/);
assert.doesNotMatch(script, /open POSIX file|close .*presentation/);
const compiled = path.join(require("node:os").tmpdir(), `ppt-formula-${process.pid}.scpt`);
const source = path.join(require("node:os").tmpdir(), `ppt-formula-${process.pid}.applescript`);
try {
  fs.writeFileSync(source, script);
  require("node:child_process").execFileSync("/usr/bin/osacompile", ["-o", compiled, source]);
} finally {
  fs.rmSync(compiled, { force: true });
  fs.rmSync(source, { force: true });
}
console.log("剪贴板接口权限模拟检查通过。");
