const http = require("http");
const fs = require("fs");
const path = require("path");

const rootDir = __dirname;
const port = Number(process.argv[2] || 8081);

const mimeTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".png", "image/png"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
  [".svg", "image/svg+xml"],
  [".webmanifest", "application/manifest+json; charset=utf-8"],
  [".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
  [".ico", "image/x-icon"]
]);

function send(response, statusCode, headers, body) {
  response.writeHead(statusCode, headers);
  response.end(body);
}

function safeResolve(requestPath) {
  const decodedPath = decodeURIComponent(requestPath.split("?")[0]);
  const normalizedPath = decodedPath === "/" ? "/index.html" : decodedPath;
  const resolvedPath = path.resolve(rootDir, "." + normalizedPath);
  if (!resolvedPath.startsWith(rootDir)) {
    return null;
  }
  return resolvedPath;
}

const server = http.createServer((request, response) => {
  const resolvedPath = safeResolve(request.url || "/");
  if (!resolvedPath) {
    send(response, 403, { "Content-Type": "text/plain; charset=utf-8" }, "Forbidden");
    return;
  }

  fs.stat(resolvedPath, (statError, stats) => {
    if (statError) {
      send(response, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Not Found");
      return;
    }

    const filePath = stats.isDirectory() ? path.join(resolvedPath, "index.html") : resolvedPath;
    fs.readFile(filePath, (readError, data) => {
      if (readError) {
        send(response, 404, { "Content-Type": "text/plain; charset=utf-8" }, "Not Found");
        return;
      }

      const extension = path.extname(filePath).toLowerCase();
      const contentType = mimeTypes.get(extension) || "application/octet-stream";
      send(response, 200, { "Content-Type": contentType, "Cache-Control": "no-store" }, data);
    });
  });
});

server.listen(port, "127.0.0.1", () => {
  console.log("Static server running at http://127.0.0.1:" + port);
});
