/**
 * Unit tests for folder browser helpers
 */

import { describe, it, expect } from 'vitest';
import { describeWebUrl } from '../../../src/ui/folderBrowserModal';

describe('describeWebUrl', () => {
	it('keeps the host and the /sites/<name> part of a library URL', () => {
		expect(describeWebUrl('https://contoso.sharepoint.com/sites/Team/Shared%20Documents')).toBe(
			'contoso.sharepoint.com/sites/Team'
		);
	});

	it('keeps the /teams/<name> part and decodes it', () => {
		expect(describeWebUrl('https://contoso.sharepoint.com/teams/Field%20Ops/Notes/Vault')).toBe(
			'contoso.sharepoint.com/teams/Field Ops'
		);
	});

	it('uses only the host for a library on the root site', () => {
		expect(describeWebUrl('https://contoso.sharepoint.com/Shared%20Documents')).toBe(
			'contoso.sharepoint.com'
		);
	});

	it('returns undefined for a missing or unparseable URL', () => {
		expect(describeWebUrl(undefined)).toBeUndefined();
		expect(describeWebUrl('not a url')).toBeUndefined();
	});
});
