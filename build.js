// Menggabungkan index.html + css + js + library Excel menjadi SATU file mandiri (bisa dibuka tanpa internet).
// Jalankan: node build.js   ->   hasil: dist/GIT_Asset_Online.html
const fs = require("fs");
const read = f => fs.readFileSync(f, "utf8");
const safe = s => s.replace(/<\/script/gi, "<\\/script").replace(/\uFFFD/g, "\\ufffd");
let html = read("index.html");
html = html.replace('<link rel="stylesheet" href="css/style.css">', () => "<style>" + read("css/style.css") + "</style>");
html = html.replace('<script src="vendor/xlsx.full.min.js"></script>', () => "<script>" + safe(read("vendor/xlsx.full.min.js")) + "</script>");
html = html.replace('<script src="js/app.js"></script>', () => "<script>" + safe(read("js/app.js")) + "</script>");
fs.mkdirSync("dist", { recursive: true });
fs.writeFileSync("dist/GIT_Asset_Online.html", html);
console.log("OK -> dist/GIT_Asset_Online.html (" + Math.round(html.length / 1024) + " KB)");
