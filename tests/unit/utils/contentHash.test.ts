/**
 * Unit tests for content hashing helper
 */

import { describe, it, expect } from 'vitest';
import { quickXorHash, sha1HexUpper } from '../../../src/utils/contentHash';
import { QUICK_XOR_HASH_VECTORS } from '../../fixtures/quickXorHashVectors';

const encode = (text: string): ArrayBuffer => new TextEncoder().encode(text).buffer;

const decodeBase64 = (value: string): Uint8Array =>
	Uint8Array.from(atob(value), (char) => char.charCodeAt(0));

describe('quickXorHash', () => {
	it.each(
		QUICK_XOR_HASH_VECTORS.map(([input, expected]) => [decodeBase64(input).length, input, expected])
	)('matches the reference vector for a %i-byte input', (_length, input, expected) => {
		expect(quickXorHash(decodeBase64(input as string))).toBe(expected);
	});

	it('gives the same result for an ArrayBuffer and a Uint8Array', () => {
		const [input, expected] = QUICK_XOR_HASH_VECTORS[QUICK_XOR_HASH_VECTORS.length - 1];
		const bytes = decodeBase64(input);
		const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
		expect(quickXorHash(buffer)).toBe(expected);
	});

	it('changes when same-length content changes', () => {
		expect(quickXorHash(encode('hello world'))).not.toBe(quickXorHash(encode('hello worle')));
	});
});

describe('sha1HexUpper', () => {
	it('computes the SHA-1 of an empty buffer as uppercase hex', async () => {
		expect(await sha1HexUpper(encode(''))).toBe('DA39A3EE5E6B4B0D3255BFEF95601890AFD80709');
	});

	it('computes a known SHA-1 vector as uppercase hex', async () => {
		expect(await sha1HexUpper(encode('abc'))).toBe('A9993E364706816ABA3E25717850C26C9CD0D89D');
	});

	it('returns undefined when WebCrypto is unavailable', async () => {
		const original = globalThis.crypto;
		try {
			Object.defineProperty(globalThis, 'crypto', { value: undefined, configurable: true });
			expect(await sha1HexUpper(encode('abc'))).toBeUndefined();
		} finally {
			Object.defineProperty(globalThis, 'crypto', { value: original, configurable: true });
		}
	});
});
