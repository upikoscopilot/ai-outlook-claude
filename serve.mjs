// Minimal HTTPS static server for the built add-in (dist/) on https://localhost:3000
// Uses the trusted localhost certificate created by "npm run certs" (office-addin-dev-certs).
// No dependencies besides Node.js.
import https from "node:https";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "dist");
const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const port = Number(process.env.PORT) || 3000;

const types = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json",
  ".xml": "application/xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".svg": "image/svg+xml",
  ".map": "application/json",
};

let options;
try {
  options = {
    key: fs.readFileSync(path.join(certDir, "localhost.key")),
    cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
  };
} catch {
  console.error(`Zertifikat fehlt in ${certDir}. Bitte zuerst "npm run certs" ausführen.`);
  process.exit(1);
}

https
  .createServer(options, (req, res) => {
    const urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
    let file = path.normalize(path.join(root, urlPath));
    if (!file.startsWith(root)) {
      res.writeHead(403).end();
      return;
    }
    if (!fs.existsSync(file) || fs.statSync(file).isDirectory()) file = path.join(root, "index.html");
    res.writeHead(200, {
      "Content-Type": types[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-cache",
    });
    fs.createReadStream(file).pipe(res);
  })
  .listen(port, "127.0.0.1", () => console.log(`Claude Mail Add-in läuft auf https://localhost:${port}/ (Fenster offen lassen)`));
