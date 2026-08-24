const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');

const PORT = 3000;
const PUBLIC_DIR = __dirname;
const CERT_DIR = path.join(os.homedir(), '.office-addin-dev-certs');

// Muat sertifikat SSL localhost
let sslOptions = {};
try {
    const keyPath = path.join(CERT_DIR, 'localhost.key');
    const certPath = path.join(CERT_DIR, 'localhost.crt');
    if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
        sslOptions = {
            key: fs.readFileSync(keyPath),
            cert: fs.readFileSync(certPath)
        };
        console.log("Sertifikat SSL dev ditemukan dan dimuat.");
    } else {
        console.warn("Peringatan: Sertifikat SSL tidak ditemukan di ~/.office-addin-dev-certs");
    }
} catch (err) {
    console.error("Gagal membaca sertifikat SSL:", err);
}

const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.xml': 'application/xml; charset=utf-8',
    '.ico': 'image/x-icon'
};

// Request Handler
const requestHandler = (req, res) => {
    // Tambahkan header CORS untuk Office Add-in
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    let reqPath = req.url.split('?')[0];
    if (reqPath === '/' || reqPath === '') {
        reqPath = '/index.html';
    }

    // Tangani permintaan icon default jika belum ada file gambar fisik
    if (reqPath.startsWith('/assets/icon-')) {
        const svgIcon = `
            <svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">
                <rect width="64" height="64" rx="8" fill="#107c41"/>
                <text x="32" y="44" font-family="Segoe UI, sans-serif" font-size="36" font-weight="bold" fill="#ffffff" text-anchor="middle">X</text>
            </svg>
        `.trim();
        res.writeHead(200, { 'Content-Type': 'image/svg+xml' });
        res.end(svgIcon);
        return;
    }

    const filePath = path.join(PUBLIC_DIR, reqPath);

    // Cegah directory traversal
    if (!filePath.startsWith(PUBLIC_DIR)) {
        res.writeHead(403);
        res.end('Forbidden');
        return;
    }

    fs.readFile(filePath, (err, data) => {
        if (err) {
            console.log(`404 Not Found: ${reqPath}`);
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('404 Not Found');
            return;
        }

        const ext = path.extname(filePath).toLowerCase();
        const contentType = MIME_TYPES[ext] || 'application/octet-stream';
        res.writeHead(200, { 'Content-Type': contentType });
        res.end(data);
    });
};

// Buat HTTPS Server
if (sslOptions.key && sslOptions.cert) {
    https.createServer(sslOptions, requestHandler).listen(PORT, () => {
        console.log(`====================================================`);
        console.log(` Excel Smart Assistant Add-In Server Aktif!`);
        console.log(` URL Task Pane: https://localhost:${PORT}/index.html`);
        console.log(`====================================================`);
    });
} else {
    // Fallback ke HTTP jika SSL belum terinstal
    http.createServer(requestHandler).listen(PORT, () => {
        console.log(`====================================================`);
        console.log(` Server Aktif pada HTTP (Perlu HTTPS untuk Excel):`);
        console.log(` http://localhost:${PORT}/index.html`);
        console.log(`====================================================`);
    });
}
