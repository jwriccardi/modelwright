// Zero-dependency HTTPS static file server for the K1 keyboard-shortcut spike.
// Serves the ./web directory on https://localhost:3100 using the office-addin-dev-certs
// localhost certificate (see setup.ps1). Node's built-in http/https/fs/path modules only.
//
// Usage: node server.js
// Stop with Ctrl+C.

const fs = require("fs");
const path = require("path");
const https = require("https");
const os = require("os");

const PORT = 3100;
const WEB_ROOT = path.join(__dirname, "web");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const CERT_PATH = path.join(CERT_DIR, "localhost.crt");
const KEY_PATH = path.join(CERT_DIR, "localhost.key");

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".ico": "image/x-icon",
};

function fail(message) {
  console.error(message);
  process.exit(1);
}

if (!fs.existsSync(CERT_PATH) || !fs.existsSync(KEY_PATH)) {
  fail(
    `Dev cert not found at ${CERT_PATH} / ${KEY_PATH}.\n` +
      "Run setup.ps1 first (it installs the office-addin-dev-certs localhost certificate)."
  );
}

const options = {
  cert: fs.readFileSync(CERT_PATH),
  key: fs.readFileSync(KEY_PATH),
};

function safeJoin(root, urlPath) {
  const decoded = decodeURIComponent(urlPath.split("?")[0]);
  const normalized = path.normalize(decoded).replace(/^(\.\.[/\\])+/, "");
  const full = path.join(root, normalized);
  if (!full.startsWith(root)) {
    return null;
  }
  return full;
}

const server = https.createServer(options, (req, res) => {
  let urlPath = req.url === "/" ? "/taskpane-a.html" : req.url;
  let filePath = safeJoin(WEB_ROOT, urlPath);

  if (!filePath) {
    res.writeHead(400).end("Bad request");
    return;
  }

  fs.stat(filePath, (err, stat) => {
    if (err || !stat.isFile()) {
      res.writeHead(404, { "Content-Type": "text/plain" }).end(`Not found: ${urlPath}`);
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || "application/octet-stream";
    res.writeHead(200, {
      "Content-Type": contentType,
      "Cache-Control": "no-store",
    });
    fs.createReadStream(filePath).pipe(res);
  });
});

server.listen(PORT, () => {
  console.log(`EMT K1 spike server listening on https://localhost:${PORT}`);
  console.log(`Serving static files from: ${WEB_ROOT}`);
  console.log("Press Ctrl+C to stop.");
});
