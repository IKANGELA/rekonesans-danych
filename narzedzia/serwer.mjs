// Prosty serwer do podglądu konwertera na własnym komputerze (Node.js).
// Uruchamiany przez podglad.bat; na GitHub Pages nie jest potrzebny.
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { join, extname, normalize } from "node:path";

const KATALOG = process.argv[2];
const PORT = Number(process.argv[3] || 8765);
const TYPY = {
    ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".mjs": "text/javascript",
    ".wasm": "application/wasm", ".json": "application/json", ".zip": "application/zip",
    ".whl": "application/zip", ".py": "text/plain; charset=utf-8", ".xml": "application/xml",
};

createServer(async (req, res) => {
    const sciezka = normalize(decodeURIComponent(new URL(req.url, "http://x").pathname))
        .replace(/^([\\/])+/, "") || "index.html";   // adres główny "/" = strona konwertera
    if (sciezka.startsWith("..")) return res.writeHead(403).end();   // tylko pliki z folderu projektu
    try {
        const dane = await readFile(join(KATALOG, sciezka));
        res.writeHead(200, { "Content-Type": TYPY[extname(sciezka)] || "application/octet-stream" });
        res.end(req.method === "HEAD" ? undefined : dane);
    } catch {
        res.writeHead(404).end();
    }
}).listen(PORT, "127.0.0.1", () => console.log(`Podgląd: http://127.0.0.1:${PORT}  (zamknij okno, aby wyłączyć)`));
