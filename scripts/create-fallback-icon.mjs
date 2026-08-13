import { deflateSync } from "node:zlib";
import { writeFileSync } from "node:fs";

const size = 1024;
const pixels = Buffer.alloc(size * size * 4);

function setPixel(x, y, color) {
  if (x < 0 || y < 0 || x >= size || y >= size) return;
  const offset = (y * size + x) * 4;
  pixels[offset] = color[0];
  pixels[offset + 1] = color[1];
  pixels[offset + 2] = color[2];
  pixels[offset + 3] = color[3] ?? 255;
}

function fillRect(x, y, width, height, color) {
  for (let row = y; row < y + height; row += 1) {
    for (let column = x; column < x + width; column += 1) setPixel(column, row, color);
  }
}

for (let y = 0; y < size; y += 1) {
  for (let x = 0; x < size; x += 1) {
    const highlight = Math.round(((x + y) / (size * 2)) * 22);
    setPixel(x, y, [36 + highlight, 87 + highlight, 229 + highlight, 255]);
  }
}

const white = [255, 255, 255, 255];
const pale = [222, 233, 255, 255];
const navy = [31, 72, 190, 255];
const left = 176;
const top = 176;
const span = 672;
const line = 30;

fillRect(left, top, span, line, white);
fillRect(left, top + span - line, span, line, white);
fillRect(left, top, line, span, white);
fillRect(left + span - line, top, line, span, white);
fillRect(left + 214, top, line, span, pale);
fillRect(left + 428, top, line, span, pale);
fillRect(left, top + 214, span, line, pale);
fillRect(left, top + 428, span, line, pale);

fillRect(407, 407, 210, 210, white);
fillRect(476, 438, 72, 148, navy);
fillRect(438, 476, 148, 72, navy);

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const name = Buffer.from(type);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(Buffer.concat([name, data])));
  return Buffer.concat([length, name, data, checksum]);
}

const raw = Buffer.alloc((size * 4 + 1) * size);
for (let row = 0; row < size; row += 1) {
  const rowOffset = row * (size * 4 + 1);
  raw[rowOffset] = 0;
  pixels.copy(raw, rowOffset + 1, row * size * 4, (row + 1) * size * 4);
}

const header = Buffer.alloc(13);
header.writeUInt32BE(size, 0);
header.writeUInt32BE(size, 4);
header[8] = 8;
header[9] = 6;

const png = Buffer.concat([
  Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
  chunk("IHDR", header),
  chunk("IDAT", deflateSync(raw, { level: 9 })),
  chunk("IEND", Buffer.alloc(0)),
]);

for (const target of ["icon.png", "splash-icon.png", "favicon.png", "android-icon-foreground.png"]) {
  writeFileSync(new URL(`../assets/images/${target}`, import.meta.url), png);
}
