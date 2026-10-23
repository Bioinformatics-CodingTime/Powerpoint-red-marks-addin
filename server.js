const fs = require("node:fs");
const https = require("node:https");
const path = require("node:path");
const os = require("node:os");
const { execFileSync } = require("node:child_process");

const root = __dirname;
const clipboardEnv = { ...process.env, LANG: "en_US.UTF-8", LC_ALL: "en_US.UTF-8" };
const files = {
  "/": ["src/commands.html", "text/html; charset=utf-8"],
  "/src/commands.html": ["src/commands.html", "text/html; charset=utf-8"],
  "/src/commands.js": ["src/commands.js", "text/javascript; charset=utf-8"],
  "/assets/arrow-16.png": ["assets/arrow-16.png", "image/png"],
  "/assets/arrow-32.png": ["assets/arrow-32.png", "image/png"],
  "/assets/arrow-80.png": ["assets/arrow-80.png", "image/png"],
  "/assets/box-16.png": ["assets/box-16.png", "image/png"],
  "/assets/box-32.png": ["assets/box-32.png", "image/png"],
  "/assets/box-80.png": ["assets/box-80.png", "image/png"],
};

function renderNativeFormula(expression) {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "ppt-markdown-"));
  const presentationPath = path.join(tempDir, "formula.pptx");
  try {
    execFileSync("pandoc", ["-f", "markdown", "-t", "pptx", "-o", presentationPath], {
      input: `$$${expression}$$\n`, timeout: 10000, maxBuffer: 1024 * 1024,
    });
    const slideXml = execFileSync("/usr/bin/unzip", ["-p", presentationPath, "ppt/slides/slide1.xml"], {
      encoding: "utf8", timeout: 3000, maxBuffer: 1024 * 1024,
    });
    if (!slideXml.includes("<m:oMath")) throw new Error("Markdown 公式未生成原生公式对象。");
    return fs.readFileSync(presentationPath).toString("base64");
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

function placeNativeFormula(targetIndex, sourceIndex) {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "ppt-markdown-"));
  const bodyPath = path.join(tempDir, "place.applescript");
  try {
    // shortcut: Office JS cannot copy native math shapes between slides; use the live object interface for that copy.
    fs.writeFileSync(bodyPath, `tell application "/Applications/Microsoft PowerPoint.app"
set targetPresentation to active presentation
set targetSlide to slide ${targetIndex} of targetPresentation
set sourceSlide to slide ${sourceIndex} of targetPresentation
set oldCount to count of shapes of targetSlide
try
  set sourceShape to shape 1 of sourceSlide
  set sourceName to name of sourceShape
  set sourceText to content of text range of text frame of sourceShape
  copy shape sourceShape
  set slide of view of active window to targetSlide
  delay 0.5
  if (slide ID of slide of view of active window) is not (slide ID of targetSlide) then error "FORMULA_TARGET_VIEW_MISMATCH"
  paste object view of active window
  delay 0.5
  if (count of shapes of targetSlide) is not oldCount + 1 then error "FORMULA_PASTE_COUNT_MISMATCH target=" & (count of shapes of targetSlide as text) & " source=" & (count of shapes of sourceSlide as text) & " view=" & (slide ID of slide of view of active window as text) & " name=" & sourceName
  set pastedShape to shape (oldCount + 1) of targetSlide
  if (content of text range of text frame of pastedShape) is not sourceText then error "FORMULA_PASTE_TEXT_MISMATCH"
  set bridgeResult to content of text range of text frame of pastedShape
  delete sourceSlide
on error errorMessage number errorNumber
  try
    repeat while (count of shapes of targetSlide) > oldCount
      delete shape (count of shapes of targetSlide) of targetSlide
    end repeat
  end try
  try
    delete sourceSlide
  end try
  error errorMessage number errorNumber
end try
return bridgeResult
end tell
`);
    return execFileSync("/usr/bin/osascript", [bodyPath], {
      encoding: "utf8", timeout: 35000,
    }).trim();
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

const server = https.createServer(
  {
    cert: fs.readFileSync(path.join(root, ".certs/localhost.crt")),
    key: fs.readFileSync(path.join(root, ".certs/localhost.key")),
  },
  (request, response) => {
    const pathname = new URL(request.url, "https://localhost").pathname;
    if (pathname === "/formula/copy") {
      if (request.method !== "POST" || request.headers.origin !== "https://localhost:3000" || !request.headers["content-type"]?.startsWith("application/json")) {
        response.writeHead(403).end("Forbidden");
        return;
      }
      const chunks = [];
      let size = 0;
      request.on("data", (chunk) => {
        size += chunk.length;
        if (size > 50 * 1024 * 1024) {
          response.writeHead(413).end("Slide too large");
          request.destroy();
        } else chunks.push(chunk);
      });
      request.on("end", () => {
        if (size > 50 * 1024 * 1024) return;
        try {
          const markdown = execFileSync("python3", [path.join(root, "extract_formula.py")], {
            input: Buffer.concat(chunks), encoding: "utf8", timeout: 15000, maxBuffer: 1024 * 1024,
          });
          response.writeHead(200, { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" }).end(markdown);
        } catch (error) {
          response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" }).end((error.stderr?.toString() || error.message).trim().split("\n").pop());
        }
      });
      return;
    }
    if (pathname === "/clipboard/read" || pathname === "/clipboard/write" || pathname === "/formula/render" || pathname === "/formula/place") {
      if (request.method !== "POST" || request.headers.origin !== "https://localhost:3000") {
        response.writeHead(403).end("Forbidden");
        return;
      }
      if (pathname === "/clipboard/read") {
        try {
          const content = execFileSync("/usr/bin/pbpaste", { encoding: "utf8", maxBuffer: 1024 * 1024, timeout: 3000, env: clipboardEnv });
          response.writeHead(200, { "Cache-Control": "no-store", "Content-Type": "text/plain; charset=utf-8" }).end(content);
        } catch (error) {
          response.writeHead(500).end("Cannot read clipboard");
        }
        return;
      }
      if (!request.headers["content-type"]?.startsWith(pathname === "/formula/place" ? "application/json" : "text/plain")) {
        response.writeHead(415).end("Wrong content type");
        return;
      }
      const chunks = [];
      let size = 0;
      request.on("data", (chunk) => {
        size += chunk.length;
        if (size > 1024 * 1024) {
          response.writeHead(413).end("Text too large");
          request.destroy();
        } else chunks.push(chunk);
      });
      request.on("end", () => {
        if (size > 1024 * 1024) return;
        try {
          const content = Buffer.concat(chunks);
          if (pathname === "/formula/render") {
            const expression = content.toString("utf8").trim();
            if (!expression || expression.length > 5000) throw new Error("公式为空或过长。");
            response.writeHead(200, { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store" }).end(renderNativeFormula(expression));
          } else if (pathname === "/formula/place") {
            const { targetIndex, sourceIndex } = JSON.parse(content.toString("utf8"));
            if (![targetIndex, sourceIndex].every((index) => Number.isSafeInteger(index) && index > 0) || sourceIndex !== targetIndex + 1) throw new Error("幻灯片定位信息无效。");
            let originalClipboard = Buffer.alloc(0);
            try { originalClipboard = execFileSync("/usr/bin/pbpaste", { maxBuffer: 1024 * 1024, timeout: 3000, env: clipboardEnv }); }
            catch { /* An empty clipboard makes pbpaste exit with code 1. */ }
            let readBack;
            try { readBack = placeNativeFormula(targetIndex, sourceIndex); }
            finally {
              try { execFileSync("/usr/bin/pbcopy", { input: originalClipboard, timeout: 3000, env: clipboardEnv }); }
              catch (error) { console.warn("Cannot restore clipboard after formula insert", error.message); }
            }
            response.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" }).end(readBack);
          } else {
            execFileSync("/usr/bin/pbcopy", { input: content, timeout: 3000, env: clipboardEnv });
            response.writeHead(204).end();
          }
        } catch (error) {
          if (pathname.startsWith("/formula/")) console.error(pathname, error.stderr?.toString().trim() || error.message);
          response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" }).end(pathname.startsWith("/formula/") ? (error.stderr?.toString().trim() || error.message) : "Cannot write clipboard");
        }
      });
      return;
    }
    const entry = files[pathname];
    if (!entry) {
      response.writeHead(404).end("Not found");
      return;
    }

    response.writeHead(200, {
      "Cache-Control": "no-store",
      "Content-Type": entry[1],
    });
    fs.createReadStream(path.join(root, entry[0])).pipe(response);
  }
);

server.listen(3000, "127.0.0.1", () => {
  console.log("PowerPoint 快捷工具加载项服务已启动：https://localhost:3000");
});
