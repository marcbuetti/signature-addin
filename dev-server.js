/*
 * Minimal HTTPS static server for the add-in's web assets on https://localhost:3000.
 * Uses the office-addin-dev-certs trusted localhost certificate (run `npm run certs` once).
 */
const fs = require("fs");
const path = require("path");
const https = require("https");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;
const ROOT = __dirname;

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".json": "application/json; charset=utf-8",
  ".xml": "text/xml; charset=utf-8",
};

function send(res, status, body, type) {
  res.writeHead(status, {
    "Content-Type": type || "text/plain; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Cache-Control": "no-store",
  });
  res.end(body);
}

(async () => {
  const options = await devCerts.getHttpsServerOptions();

  https
    .createServer(options, (req, res) => {
      try {
        console.log(new Date().toISOString(), req.method, req.url);
        const urlPath = decodeURIComponent(req.url.split("?")[0]);
        let rel = urlPath === "/" ? "/index.html" : urlPath;
        const filePath = path.join(ROOT, rel);

        // Prevent path traversal outside the project root.
        if (!filePath.startsWith(ROOT)) {
          return send(res, 403, "Forbidden");
        }
        if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
          return send(res, 404, "Not found: " + rel);
        }
        const ext = path.extname(filePath).toLowerCase();
        send(res, 200, fs.readFileSync(filePath), TYPES[ext]);
      } catch (e) {
        send(res, 500, "Server error: " + e.message);
      }
    })
    .listen(PORT, () => {
      console.log(`Add-in dev server running at https://localhost:${PORT}/`);
      console.log(`  runtime:  https://localhost:${PORT}/commands/commands.html`);
      console.log(`  icon:     https://localhost:${PORT}/assets/icon-32.png`);
    });
})();
