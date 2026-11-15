// Server jaringan lokal (tanpa instalasi library). Jalankan: node server.js
// Perangkat lain di Wi-Fi/LAN yang sama membuka salah satu alamat yang tampil (mis. http://192.168.1.10:8080/).
const http = require("http"), fs = require("fs"), path = require("path"), os = require("os");
const PORT = process.env.PORT || 8080, ROOT = __dirname;
const TIPE = { ".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".json": "application/json", ".webmanifest": "application/manifest+json", ".png": "image/png" };
http.createServer((req, res) => {
  let u = decodeURIComponent(req.url.split("?")[0]); if (u.endsWith("/")) u += "index.html";
  const f = path.normalize(path.join(ROOT, u));
  if (!f.startsWith(ROOT) || /(^|[\\/])\.|node_modules|\.xlsx?$/i.test(path.relative(ROOT, f))) { res.writeHead(403); return res.end("Dilarang"); }
  fs.readFile(f, (e, d) => { if (e) { res.writeHead(404); return res.end("Tidak ditemukan"); } res.writeHead(200, { "Content-Type": TIPE[path.extname(f)] || "application/octet-stream", "Cache-Control": "no-cache" }); res.end(d); });
}).listen(PORT, "0.0.0.0", () => {
  console.log("GIT Asset Online berjalan. Buka dari perangkat di jaringan yang sama:");
  Object.values(os.networkInterfaces()).flat().filter(i => i.family === "IPv4" && !i.internal).forEach(i => console.log("  http://" + i.address + ":" + PORT + "/"));
  console.log("  http://localhost:" + PORT + "/  (komputer ini)\nTekan Ctrl+C untuk berhenti. Bila tidak bisa dibuka dari HP, izinkan Node.js di Windows Firewall (jaringan Private).");
});
