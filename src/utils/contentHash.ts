/**
 * Content hashing for local↔remote comparison.
 *
 * OneDrive Personal returns a `sha1Hash` in the driveItem hashes facet
 * (uppercase hex). OneDrive for Business and SharePoint generally do not:
 * `quickXorHash` is the only hash Microsoft guarantees on both. Computing
 * either locally lets us verify that a same-size local file really has the
 * same content as the remote copy, instead of relying on the size
 * heuristic alone.
 */

const QUICK_XOR_WIDTH_BITS = 160;
const QUICK_XOR_SHIFT = 11;

/**
 * Compute Microsoft's QuickXorHash of a buffer, base64-encoded (the format
 * OneDrive uses for `quickXorHash`).
 *
 * The algorithm XORs byte i into a 160-bit circular register at bit offset
 * (i * 11) mod 160, then XORs the 64-bit little-endian data length into the
 * last 8 bytes. Because the offset repeats every 160 bytes, bytes are first
 * folded into 160 buckets (one per offset), so placing them is a fixed cost
 * however large the file is.
 *
 * Reference: https://learn.microsoft.com/onedrive/developer/code-snippets/quickxorhash
 */
export function quickXorHash(data: ArrayBuffer | Uint8Array): string {
	const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

	const buckets = new Uint8Array(QUICK_XOR_WIDTH_BITS);
	for (let i = 0; i < bytes.length; i++) {
		buckets[i % QUICK_XOR_WIDTH_BITS] ^= bytes[i];
	}

	const register = new Uint8Array(QUICK_XOR_WIDTH_BITS / 8);
	for (let k = 0; k < QUICK_XOR_WIDTH_BITS; k++) {
		const value = buckets[k];
		if (value === 0) continue;
		const offset = (k * QUICK_XOR_SHIFT) % QUICK_XOR_WIDTH_BITS;
		for (let bit = 0; bit < 8; bit++) {
			if ((value >> bit) & 1) {
				const position = (offset + bit) % QUICK_XOR_WIDTH_BITS;
				register[position >> 3] ^= 1 << (position & 7);
			}
		}
	}

	// XOR the length (64-bit little-endian) into the final 8 bytes.
	const length = bytes.length;
	const low = length >>> 0;
	const high = Math.floor(length / 0x100000000) >>> 0;
	const lengthStart = register.length - 8;
	for (let i = 0; i < 4; i++) {
		register[lengthStart + i] ^= (low >>> (8 * i)) & 0xff;
		register[lengthStart + 4 + i] ^= (high >>> (8 * i)) & 0xff;
	}

	let binary = '';
	for (const byte of register) binary += String.fromCharCode(byte);
	return btoa(binary);
}

/**
 * Compute the SHA-1 of a buffer as an uppercase hex string (the format
 * OneDrive uses for `sha1Hash`). Returns undefined when WebCrypto is
 * unavailable so callers can fall back to the size heuristic.
 */
export async function sha1HexUpper(data: ArrayBuffer): Promise<string | undefined> {
	if (typeof crypto === 'undefined' || !crypto.subtle) {
		return undefined;
	}
	const digest = await crypto.subtle.digest('SHA-1', data);
	return Array.from(new Uint8Array(digest))
		.map((byte) => byte.toString(16).padStart(2, '0'))
		.join('')
		.toUpperCase();
}
