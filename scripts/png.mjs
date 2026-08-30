import { deflateSync } from "node:zlib";

const CRC_TABLE = (() => {
	const table = new Int32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		table[n] = c;
	}
	return table;
})();

function crc32(buf) {
	let c = -1;
	for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
	return (c ^ -1) >>> 0;
}

function chunk(type, data) {
	const out = Buffer.alloc(data.length + 12);
	out.writeUInt32BE(data.length, 0);
	out.write(type, 4, "ascii");
	data.copy(out, 8);
	const crcInput = Buffer.concat([Buffer.from(type, "ascii"), data]);
	out.writeUInt32BE(crc32(crcInput), data.length + 8);
	return out;
}

/**
 * Encode raw RGB pixels as a PNG. Just enough of the format to produce a real
 * image for the fixture deck without pulling in an encoder dependency.
 */
export function encodePng(width, height, pixelAt) {
	const raw = Buffer.alloc(height * (width * 3 + 1));
	let p = 0;
	for (let y = 0; y < height; y++) {
		raw[p++] = 0; // filter: none
		for (let x = 0; x < width; x++) {
			const [r, g, b] = pixelAt(x, y);
			raw[p++] = r;
			raw[p++] = g;
			raw[p++] = b;
		}
	}

	const ihdr = Buffer.alloc(13);
	ihdr.writeUInt32BE(width, 0);
	ihdr.writeUInt32BE(height, 4);
	ihdr[8] = 8; // bit depth
	ihdr[9] = 2; // colour type: truecolour
	ihdr[10] = 0;
	ihdr[11] = 0;
	ihdr[12] = 0;

	return Buffer.concat([
		Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
		chunk("IHDR", ihdr),
		chunk("IDAT", deflateSync(raw, { level: 9 })),
		chunk("IEND", Buffer.alloc(0)),
	]);
}
