'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { deflateSync } = require('node:zlib');

function crc32(data) {
  let value = 0xffffffff;
  for (const byte of data) {
    value ^= byte;
    for (let bit = 0; bit < 8; bit++) value = value >>> 1 ^ (value & 1 ? 0xedb88320 : 0);
  }
  return (value ^ 0xffffffff) >>> 0;
}
function chunk(name, data) {
  const type = Buffer.from(name), length = Buffer.alloc(4), crc = Buffer.alloc(4);
  length.writeUInt32BE(data.length); crc.writeUInt32BE(crc32(Buffer.concat([type, data])));
  return Buffer.concat([length, type, data, crc]);
}
function makePNG(size) {
  const pixels = Buffer.alloc((size * 4 + 1) * size);
  for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
    const u = x / size, v = y / size, offset = y * (size * 4 + 1) + 1 + x * 4;
    const inset = 0.045, radius = 0.16;
    const cx = Math.max(inset + radius, Math.min(1 - inset - radius, u));
    const cy = Math.max(inset + radius, Math.min(1 - inset - radius, v));
    const inside = (u - cx) ** 2 + (v - cy) ** 2 <= radius ** 2;
    if (!inside) continue;
    let color = [16, Math.round(124 + 15 * (1 - v)), 65, 255];
    const sheet = u > 0.235 && u < 0.815 && v > 0.19 && v < 0.81;
    if (sheet) color = [234, 249, 241, 255];
    if (sheet && v < 0.34) color = [143, 216, 176, 255];
    const grid = (u > 0.405 && u < 0.435 || u > 0.605 && u < 0.635) && v > 0.34 && v < 0.77 ||
      (v > 0.475 && v < 0.50 || v > 0.625 && v < 0.65) && u > 0.27 && u < 0.78;
    if (sheet && grid) color = [39, 137, 82, 255];
    pixels.set(color, offset);
  }
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0); header.writeUInt32BE(size, 4); header[8] = 8; header[9] = 6;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', header), chunk('IDAT', deflateSync(pixels)), chunk('IEND', Buffer.alloc(0))]);
}
async function createIcon(directory) {
  await fs.mkdir(directory, { recursive: true });
  const images = [32, 64, 128, 256].map(size => ({ size, bytes: makePNG(size) }));
  const header = Buffer.alloc(6 + images.length * 16);
  header.writeUInt16LE(1, 2); header.writeUInt16LE(images.length, 4);
  let offset = header.length;
  images.forEach(({ size, bytes }, index) => {
    const at = 6 + index * 16;
    header[at] = size === 256 ? 0 : size; header[at + 1] = size === 256 ? 0 : size;
    header.writeUInt16LE(1, at + 4); header.writeUInt16LE(32, at + 6);
    header.writeUInt32LE(bytes.length, at + 8); header.writeUInt32LE(offset, at + 12);
    offset += bytes.length;
  });
  const icon = path.join(directory, 'app.ico');
  await fs.writeFile(icon, Buffer.concat([header, ...images.map(image => image.bytes)]));
  await fs.writeFile(path.join(directory, 'app.png'), images.at(-1).bytes);
  return icon;
}

async function brandExecutable(filename, iconPath, version, root) {
  const ResEdit = require(require.resolve('resedit', { paths: [root, path.dirname(root)] }));
  const exe = ResEdit.NtExecutable.from(await fs.readFile(filename), { ignoreCert: true });
  const hash = bytes => createHash('sha256').update(Buffer.from(bytes)).digest('hex');
  const code = exe.getAllSections().filter(section => section.info.characteristics & 0x20000000)
    .map(section => ({ name: section.info.name, hash: hash(section.data) }));
  const resources = ResEdit.NtExecutableResource.from(exe);
  const icon = ResEdit.Data.IconFile.from(await fs.readFile(iconPath));
  for (const group of ResEdit.Resource.IconGroupEntry.fromEntries(resources.entries)) {
    ResEdit.Resource.IconGroupEntry.replaceIconsForResource(resources.entries, group.id, group.lang, icon.icons.map(item => item.data));
  }
  for (const info of ResEdit.Resource.VersionInfo.fromEntries(resources.entries)) {
    info.setFileVersion(version + '.0'); info.setProductVersion(version + '.0');
    for (const language of info.getAllLanguagesForStringValues()) info.setStringValues(language, {
      CompanyName: 'Sheet Studio', FileDescription: 'Sheet Studio', ProductName: 'Sheet Studio',
      OriginalFilename: 'Sheet Studio.exe', InternalName: 'Sheet Studio', FileVersion: version, ProductVersion: version,
      LegalCopyright: 'Independent spreadsheet application.',
    });
    info.outputToResourceEntries(resources.entries);
  }
  resources.entries = resources.entries.filter(entry => entry.type !== 'INTEGRITY');
  resources.outputResource(exe);
  const output = Buffer.from(exe.generate());
  const verified = ResEdit.NtExecutable.from(output);
  for (const section of code) {
    const after = verified.getAllSections().find(item => item.info.name === section.name);
    if (!after || hash(after.data) !== section.hash) throw new Error('Executable code changed while updating product metadata.');
  }
  await fs.writeFile(filename, output);
}

module.exports = { createIcon, brandExecutable };
