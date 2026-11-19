/**
 * Modal for browsing and selecting OneDrive folders
 */

import { App, Modal, Setting } from 'obsidian';
import { OneDriveItem, SharePointFolderTarget, SharePointLibrary } from '../types';
import { RootFolderWarningModal } from './modals';
import { t } from '../i18n';

export interface FolderSelection {
	path: string;
	name: string;
	isShared: boolean;
	driveId?: string;
	itemId?: string;
	relativePathInShared?: string;
	/** itemId is the root of its drive: a whole SharePoint document library */
	isDriveRoot?: boolean;
	/** Browser URL of the SharePoint library or folder the selection is in */
	webUrl?: string;
}

type FolderListFn = (
	path: string,
	sharedDriveId?: string,
	sharedItemId?: string,
	relativePathInShared?: string
) => Promise<OneDriveItem[]>;

/**
 * SharePoint discovery for work and school accounts. When supplied, the top
 * level of the picker offers library search and "open a link"; choosing one
 * browses into it like any other folder on another drive.
 */
export interface SharePointPickerApi {
	searchLibraries(query: string): Promise<SharePointLibrary[]>;
	resolveLink(url: string): Promise<SharePointFolderTarget>;
	getLibraryRoot(driveId: string): Promise<{ id: string; name: string }>;
}

export interface FolderBrowserOptions {
	/** Label for the root breadcrumb (default: "OneDrive") */
	rootLabel?: string;
	/** Show a confirmation warning when the drive root is selected (default: false) */
	warnOnRootSelect?: boolean;
	/** Offer SharePoint libraries at the top level (work and school accounts) */
	sharePoint?: SharePointPickerApi;
}

/**
 * Modal that lets the user browse OneDrive folders and select one
 */
export class FolderBrowserModal extends Modal {
	private currentPath: string[] = [];
	private onSelect: (selection: FolderSelection) => void;
	private listFolders: FolderListFn;
	private contentEl_body: HTMLElement;
	// Incremented on every navigation; a load whose number is no longer current
	// stops drawing, so the newest navigation always wins.
	private loadGeneration = 0;
	private rootLabel: string;
	private warnOnRootSelect: boolean;

	// Track the shared folder info if the user navigated into one
	private sharedDriveId?: string;
	private sharedItemId?: string;
	private sharedAtDepth?: number; // depth at which the shared folder was entered
	// Set while browsing a SharePoint library or folder opened from the
	// SharePoint section (as opposed to a shortcut in the user's own drive)
	private sharePointTarget?: { isDriveRoot: boolean; webUrl?: string };
	private sharePoint?: SharePointPickerApi;

	constructor(
		app: App,
		listFolders: FolderListFn,
		onSelect: (selection: FolderSelection) => void,
		initialPath?: string,
		options?: FolderBrowserOptions
	) {
		super(app);
		this.listFolders = listFolders;
		this.onSelect = onSelect;
		this.rootLabel = options?.rootLabel ?? 'OneDrive';
		this.warnOnRootSelect = options?.warnOnRootSelect ?? false;
		this.sharePoint = options?.sharePoint;

		if (initialPath) {
			this.currentPath = initialPath.split('/').filter((s) => s.length > 0);
		}
	}

	onOpen() {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.addClass('onedrive-folder-browser');

		contentEl.createEl('h3', { text: t('folderBrowser.title') });

		this.contentEl_body = contentEl.createDiv({
			cls: 'folder-browser-body onedrive-sync-folder-browser-body',
		});

		void this.loadFolder();
	}

	onClose() {
		this.contentEl.empty();
	}

	private get currentPathStr(): string {
		return this.currentPath.length > 0 ? this.currentPath.join('/') : '';
	}

	/**
	 * Get the relative path within the shared folder (segments after sharedAtDepth + 1)
	 */
	private get relativePathInShared(): string {
		if (this.sharedAtDepth === undefined) return '';
		const segments = this.currentPath.slice(this.sharedAtDepth + 1);
		return segments.join('/');
	}

	private async loadFolder() {
		const generation = ++this.loadGeneration;

		const body = this.contentEl_body;
		body.empty();

		// Breadcrumb
		const breadcrumb = body.createDiv({
			cls: 'folder-breadcrumb onedrive-sync-folder-breadcrumb',
		});

		// Root link
		const rootLink = breadcrumb.createSpan({
			text: `📁 ${this.rootLabel}`,
			cls: 'onedrive-sync-folder-breadcrumb-link',
		});
		rootLink.onclick = () => {
			this.currentPath = [];
			this.sharedDriveId = undefined;
			this.sharedItemId = undefined;
			this.sharedAtDepth = undefined;
			this.sharePointTarget = undefined;
			void this.loadFolder();
		};

		for (let i = 0; i < this.currentPath.length; i++) {
			breadcrumb.createSpan({ text: ' / ' });
			const seg = breadcrumb.createSpan({ text: this.currentPath[i] });
			if (i < this.currentPath.length - 1) {
				seg.addClass('onedrive-sync-folder-breadcrumb-link');
				const depth = i;
				seg.onclick = () => {
					this.currentPath = this.currentPath.slice(0, depth + 1);
					if (this.sharedAtDepth !== undefined && depth < this.sharedAtDepth) {
						this.sharedDriveId = undefined;
						this.sharedItemId = undefined;
						this.sharedAtDepth = undefined;
						this.sharePointTarget = undefined;
					}
					void this.loadFolder();
				};
			}
		}

		// Select button for the current folder. This is also shown at the drive
		// root (empty path) so the user can sync their entire OneDrive if desired.
		// The root can only ever be the user's own drive — shared folders are
		// always entered at depth >= 1 — so root selection is never "shared".
		{
			const isRoot = this.currentPath.length === 0;
			const selectRow = body.createDiv({ cls: 'onedrive-sync-folder-browser-select-row' });
			new Setting(selectRow)
				.setName(
					isRoot
						? t('folderBrowser.selectRoot', { label: this.rootLabel })
						: t('folderBrowser.selectCurrent', { path: this.currentPath.join('/') })
				)
				.addButton((btn) =>
					btn
						.setButtonText(t('folderBrowser.useThisFolder'))
						.setCta()
						.onClick(() => {
							const isShared = !!(this.sharedDriveId && this.sharedItemId);
							const commit = () => {
								this.onSelect({
									path: `/${this.currentPath.join('/')}`,
									name: isRoot
										? this.rootLabel
										: this.currentPath[this.sharedAtDepth ?? this.currentPath.length - 1],
									isShared,
									driveId: this.sharedDriveId,
									itemId: this.sharedItemId,
									relativePathInShared: isShared ? this.relativePathInShared : undefined,
									isDriveRoot: isShared ? this.sharePointTarget?.isDriveRoot : undefined,
									webUrl: isShared ? this.sharePointTarget?.webUrl : undefined,
								});
								this.close();
							};

							// Syncing the whole drive root is a footgun — confirm first
							// when the caller opted in (full-drive picker only).
							if (isRoot && this.warnOnRootSelect) {
								new RootFolderWarningModal(this.app, commit).open();
							} else {
								commit();
							}
						})
				);
		}

		// SharePoint libraries are reachable only from the top of the user's
		// own drive, and only for work and school accounts.
		if (this.sharePoint && this.currentPath.length === 0 && !this.sharedDriveId) {
			this.renderSharePointSection(body, this.sharePoint);
		}

		// Loading indicator
		const loadingEl = body.createDiv({
			text: t('folderBrowser.loading'),
			cls: 'onedrive-sync-folder-browser-note',
		});

		try {
			let folders: OneDriveItem[];

			if (this.sharedDriveId && this.sharedItemId) {
				// Inside a shared folder — use remote drive API
				folders = await this.listFolders(
					this.currentPathStr,
					this.sharedDriveId,
					this.sharedItemId,
					this.relativePathInShared
				);
			} else {
				// User's own drive
				folders = await this.listFolders(this.currentPathStr);
			}

			if (generation !== this.loadGeneration) return; // superseded by a newer navigation
			loadingEl.remove();

			if (folders.length === 0) {
				body.createDiv({
					text: t('folderBrowser.noSubfolders'),
					cls: 'onedrive-sync-folder-browser-note onedrive-sync-folder-browser-empty',
				});
			}

			for (const folder of folders) {
				const isShared = !!folder.remoteItem;
				const name = folder.name;
				const childCount = folder.folder?.childCount ?? folder.remoteItem?.folder?.childCount ?? 0;

				const row = body.createDiv({
					cls: 'folder-row onedrive-sync-folder-row',
				});

				const icon = isShared ? '🔗' : '📁';
				row.createSpan({ text: `${icon} ${name}` });

				const meta = row.createSpan({ cls: 'onedrive-sync-folder-row-meta' });
				const parts: string[] = [];
				if (isShared) parts.push(t('folderBrowser.shared'));
				if (childCount > 0) {
					parts.push(
						t(childCount === 1 ? 'folderBrowser.item' : 'folderBrowser.items', {
							count: childCount,
						})
					);
				}
				meta.textContent = parts.join(' · ');

				row.onclick = () => {
					// If this is a shared/remote item, capture its drive info
					if (isShared && folder.remoteItem) {
						this.sharedDriveId = folder.remoteItem.parentReference.driveId;
						this.sharedItemId = folder.remoteItem.id;
						this.sharedAtDepth = this.currentPath.length;
						this.sharePointTarget = undefined;
					}
					this.currentPath.push(name);
					void this.loadFolder();
				};
			}
		} catch (error) {
			if (generation !== this.loadGeneration) return;
			loadingEl.remove();
			body.createDiv({
				text: t('folderBrowser.loadError', {
					message: error instanceof Error ? error.message : t('folderBrowser.unknownError'),
				}),
				cls: 'onedrive-sync-folder-browser-error',
			});
		}
	}

	/**
	 * Browse into a SharePoint library or folder. It becomes a "shared" root
	 * at depth 0, so the existing shared-folder browsing and selection apply
	 * unchanged; anything selected below it is a relative path inside it.
	 */
	private openSharePointTarget(target: {
		driveId: string;
		itemId: string;
		name: string;
		isDriveRoot: boolean;
		webUrl?: string;
	}): void {
		this.currentPath = [target.name];
		this.sharedDriveId = target.driveId;
		this.sharedItemId = target.itemId;
		this.sharedAtDepth = 0;
		this.sharePointTarget = { isDriveRoot: target.isDriveRoot, webUrl: target.webUrl };
		void this.loadFolder();
	}

	private renderSharePointSection(body: HTMLElement, api: SharePointPickerApi): void {
		const section = body.createDiv({ cls: 'onedrive-sync-sharepoint-section' });
		new Setting(section).setName(t('folderBrowser.sharePoint.heading')).setHeading();

		// `status` and `results` are created below the search box (in display
		// order); these handlers only run after that, on user input.
		const showError = (error: unknown) => {
			status.setText(
				t('folderBrowser.sharePoint.error', {
					message: error instanceof Error ? error.message : t('folderBrowser.unknownError'),
				})
			);
			status.addClass('onedrive-sync-folder-browser-error');
		};
		const clearStatus = (text = '') => {
			status.setText(text);
			status.removeClass('onedrive-sync-folder-browser-error');
		};

		// Find a library by name
		let query = '';
		const runSearch = async () => {
			if (!query.trim()) return;
			results.empty();
			clearStatus(t('folderBrowser.sharePoint.searching'));
			try {
				const libraries = await api.searchLibraries(query);
				clearStatus(
					libraries.length === 0
						? t('folderBrowser.sharePoint.noResults', { query: query.trim() })
						: ''
				);
				for (const library of libraries) {
					const row = results.createDiv({ cls: 'folder-row onedrive-sync-folder-row' });
					row.createSpan({ text: `🏢 ${library.name}` });
					row.createSpan({
						text: describeWebUrl(library.webUrl) ?? t('folderBrowser.sharePoint.library'),
						cls: 'onedrive-sync-folder-row-meta',
					});
					const openLibrary = async () => {
						clearStatus(t('folderBrowser.sharePoint.opening'));
						try {
							const root = await api.getLibraryRoot(library.driveId);
							this.openSharePointTarget({
								driveId: library.driveId,
								itemId: root.id,
								name: library.name,
								isDriveRoot: true,
								webUrl: library.webUrl,
							});
						} catch (error) {
							showError(error);
						}
					};
					row.onclick = () => void openLibrary();
				}
			} catch (error) {
				showError(error);
			}
		};
		new Setting(section)
			.setName(t('folderBrowser.sharePoint.searchName'))
			.setDesc(t('folderBrowser.sharePoint.searchDesc'))
			.addText((text) => {
				text.setPlaceholder(t('folderBrowser.sharePoint.searchPlaceholder')).onChange((value) => {
					query = value;
				});
				text.inputEl.addEventListener('keydown', (event) => {
					if (event.key === 'Enter') void runSearch();
				});
			})
			.addButton((btn) =>
				btn
					.setButtonText(t('folderBrowser.sharePoint.searchButton'))
					.onClick(() => void runSearch())
			);

		// Search status and results sit under the search box; link errors share
		// the status line.
		const status = section.createDiv({ cls: 'onedrive-sync-folder-browser-note' });
		const results = section.createDiv({ cls: 'onedrive-sync-sharepoint-results' });

		// Open a pasted link
		let link = '';
		const openLink = async () => {
			if (!link.trim()) return;
			clearStatus(t('folderBrowser.sharePoint.opening'));
			try {
				const target = await api.resolveLink(link);
				this.openSharePointTarget(target);
			} catch (error) {
				showError(error);
			}
		};
		new Setting(section)
			.setName(t('folderBrowser.sharePoint.linkName'))
			.setDesc(t('folderBrowser.sharePoint.linkDesc'))
			.addText((text) => {
				text.setPlaceholder(t('folderBrowser.sharePoint.linkPlaceholder')).onChange((value) => {
					link = value;
				});
				text.inputEl.addEventListener('keydown', (event) => {
					if (event.key === 'Enter') void openLink();
				});
			})
			.addButton((btn) =>
				btn.setButtonText(t('folderBrowser.sharePoint.linkButton')).onClick(() => void openLink())
			);

		new Setting(body).setName(t('folderBrowser.sharePoint.myFilesHeading')).setHeading();
	}
}

/**
 * Short "where is this" label for a SharePoint URL, e.g.
 * "contoso.sharepoint.com/sites/Team", so same-named libraries
 * ("Documents" exists on every site) can be told apart.
 */
export function describeWebUrl(webUrl?: string): string | undefined {
	if (!webUrl) return undefined;
	try {
		const url = new URL(webUrl);
		const segments = url.pathname
			.split('/')
			.filter(Boolean)
			.map((segment) => decodeURIComponent(segment));
		// Keep the site part (/sites/<name> or /teams/<name>), not the library path.
		const sitePrefix = ['sites', 'teams'].includes(segments[0]) ? segments.slice(0, 2) : [];
		return [url.host, ...sitePrefix].join('/');
	} catch {
		return undefined;
	}
}
