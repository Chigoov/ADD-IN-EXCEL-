const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

// CRC32 table & function
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    crcTable[n] = c >>> 0;
}

function crc32(buf) {
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) {
        crc = crcTable[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
}

function makeChunk(type, data) {
    const len = data.length;
    const buf = Buffer.alloc(4 + 4 + len + 4);
    buf.writeUInt32BE(len, 0);
    buf.write(type, 4, 4, 'ascii');
    data.copy(buf, 8);
    const crc = crc32(buf.subarray(4, 8 + len));
    buf.writeUInt32BE(crc, 8 + len);
    return buf;
}

function createPng(width, height, r, g, b, a = 255) {
    // PNG signature
    const signature = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

    // IHDR
    const ihdrData = Buffer.alloc(13);
    ihdrData.writeUInt32BE(width, 0);
    ihdrData.writeUInt32BE(height, 4);
    ihdrData.writeUInt8(8, 8); // 8-bit depth
    ihdrData.writeUInt8(6, 9); // RGBA
    ihdrData.writeUInt8(0, 10);
    ihdrData.writeUInt8(0, 11);
    ihdrData.writeUInt8(0, 12);
    const ihdrChunk = makeChunk('IHDR', ihdrData);

    // IDAT
    const rowSize = 1 + width * 4;
    const rawData = Buffer.alloc(rowSize * height);
    for (let y = 0; y < height; y++) {
        const rowOffset = y * rowSize;
        rawData.writeUInt8(0, rowOffset); // Filter type 0 (None)
        for (let x = 0; x < width; x++) {
            const pixelOffset = rowOffset + 1 + x * 4;
            // Draw a green box with white center border
            const isBorder = (x < 2 || x >= width - 2 || y < 2 || y >= height - 2);
            if (isBorder) {
                rawData.writeUInt8(16, pixelOffset);     // #10
                rawData.writeUInt8(124, pixelOffset + 1); // #7c
                rawData.writeUInt8(65, pixelOffset + 2);  // #41
                rawData.writeUInt8(255, pixelOffset + 3);
            } else {
                rawData.writeUInt8(r, pixelOffset);
                rawData.writeUInt8(g, pixelOffset + 1);
                rawData.writeUInt8(b, pixelOffset + 2);
                rawData.writeUInt8(a, pixelOffset + 3);
            }
        }
    }
    const compressed = zlib.deflateSync(rawData);
    const idatChunk = makeChunk('IDAT', compressed);

    // IEND
    const iendChunk = makeChunk('IEND', Buffer.alloc(0));

    return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

const assetsDir = path.join(__dirname, 'assets');
if (!fs.existsSync(assetsDir)) {
    fs.mkdirSync(assetsDir, { recursive: true });
}

// Generate Excel Green Icons: #107C41 (16, 124, 65)
fs.writeFileSync(path.join(assetsDir, 'icon-16.png'), createPng(16, 16, 16, 124, 65));
fs.writeFileSync(path.join(assetsDir, 'icon-32.png'), createPng(32, 32, 16, 124, 65));
fs.writeFileSync(path.join(assetsDir, 'icon-80.png'), createPng(80, 80, 16, 124, 65));

console.log('PNG Icons successfully generated in assets/ folder!');
