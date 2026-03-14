import http from "http";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const publicDir = path.join(rootDir, "src");
const dataDir = path.join(rootDir, "data");
const dataFile = path.join(dataDir, "records.json");
const excelTemplateCandidates = [
  path.join(rootDir, "月次日常点検 2026.xlsx"),
  path.join(rootDir, "月次日常点検 2026 (1).xlsx")
];

if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
if (!fs.existsSync(dataFile)) {
  fs.writeFileSync(dataFile, JSON.stringify({ records: {} }, null, 2), "utf8");
}

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
};

function sendJson(res, status, payload) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += chunk;
      if (body.length > 1_000_000) {
        reject(new Error("Payload too large"));
        req.destroy();
      }
    });
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(new Error("Invalid JSON"));
      }
    });
  });
}

function readStore() {
  return JSON.parse(fs.readFileSync(dataFile, "utf8"));
}

function writeStore(store) {
  fs.writeFileSync(dataFile, JSON.stringify(store, null, 2), "utf8");
}

function getRecordKey(month, vehicle, driver) {
  return `${month}__${vehicle}__${driver}`;
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

  if (url.pathname === "/api/record" && req.method === "GET") {
    const month = url.searchParams.get("month") || "2026-01";
    const vehicle = url.searchParams.get("vehicle") || "";
    const driver = url.searchParams.get("driver") || "";
    const store = readStore();
    const key = getRecordKey(month, vehicle, driver);
    const record = store.records[key] || null;
    return sendJson(res, 200, { key, record });
  }

  if (url.pathname === "/api/record" && req.method === "POST") {
    try {
      const payload = await parseBody(req);
      const { month, vehicle, driver, record } = payload;
      if (!month || !vehicle || !driver || !record) {
        return sendJson(res, 400, { error: "month, vehicle, driver, record are required" });
      }
      const store = readStore();
      const key = getRecordKey(month, vehicle, driver);
      store.records[key] = {
        ...record,
        updatedAt: new Date().toISOString()
      };
      writeStore(store);
      return sendJson(res, 200, { ok: true, key });
    } catch (err) {
      return sendJson(res, 400, { error: err.message });
    }
  }

  if (url.pathname === "/api/excel-template" && req.method === "GET") {
    const templatePath = excelTemplateCandidates.find((candidate) => fs.existsSync(candidate));
    if (!templatePath) {
      res.writeHead(404);
      return res.end("Template Not Found");
    }

    fs.readFile(templatePath, (err, data) => {
      if (err) {
        res.writeHead(500);
        return res.end("Failed to read template");
      }

      res.writeHead(200, {
        "Content-Type": contentTypes[".xlsx"],
        "Content-Disposition": `inline; filename*=UTF-8''${encodeURIComponent(path.basename(templatePath))}`
      });
      res.end(data);
    });
    return;
  }

  let filePath = url.pathname === "/" ? "/index.html" : url.pathname;
  filePath = path.normalize(filePath).replace(/^\\+/, "");
  const absPath = path.join(publicDir, filePath);

  if (!absPath.startsWith(publicDir)) {
    res.writeHead(403);
    return res.end("Forbidden");
  }

  fs.readFile(absPath, (err, data) => {
    if (err) {
      res.writeHead(404);
      return res.end("Not Found");
    }
    const ext = path.extname(absPath);
    res.writeHead(200, { "Content-Type": contentTypes[ext] || "text/plain; charset=utf-8" });
    res.end(data);
  });
});

const port = process.env.PORT || 3000;
server.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
