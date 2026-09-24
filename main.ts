import { App, Plugin, MarkdownPostProcessorContext, TFile } from 'obsidian';
import * as XLSX from 'xlsx';

interface SpreadsheetRange {
	sheet?: string;
	startCell: string;
	endCell: string;
}

interface RenderOptions {
	mode: 'stacked' | 'tabbed';
	formatting: boolean;
}

interface ParsedSpec {
	ranges: SpreadsheetRange[];
	options: RenderOptions;
}

interface WorksheetHandle {
	workbook: XLSX.WorkBook;
	file: TFile;
	filePath: string;
	filename: string;
}

export default class SpreadsheetSyncPlugin extends Plugin {
	async onload() {
		this.registerMarkdownCodeBlockProcessor('spreadsheet', async (source, el, ctx) => {
			try {
				await this.renderSpreadsheetBlock(source, el, ctx);
			} catch (error) {
				el.createEl('div', { text: `Error: ${error.message}`, cls: 'spreadsheet-error' });
			}
		});
	}

	private parseRange(rangeStr: string): SpreadsheetRange {
		const match = rangeStr.match(/(?:([^!]+)!)?([A-Z]+\d+):([A-Z]+\d+)/);
		if (!match) {
			throw new Error('Invalid range format. Expected format: [SheetName!]A1:B2');
		}
		return {
			sheet: match[1]?.trim(),
			startCell: match[2],
			endCell: match[3],
		};
	}

	private parseSpec(innerArg: string): ParsedSpec {
		// Split on first ';' — ranges part + options part.
		const semiIdx = innerArg.indexOf(';');
		const rangesPart = (semiIdx >= 0 ? innerArg.substring(0, semiIdx) : innerArg).trim();
		const optionsPart = (semiIdx >= 0 ? innerArg.substring(semiIdx + 1) : '').trim();

		// Ranges separated by ','.
		const rangeStrs = rangesPart.split(',').map(s => s.trim()).filter(Boolean);
		if (rangeStrs.length === 0) {
			throw new Error('No ranges provided');
		}
		const ranges = rangeStrs.map(s => this.parseRange(s));

		// Options: key=value pairs, ',' separated (mirrors the ranges separator).
		const options: RenderOptions = { mode: 'stacked', formatting: true };
		if (optionsPart) {
			for (const opt of optionsPart.split(',').map(s => s.trim()).filter(Boolean)) {
				const eqIdx = opt.indexOf('=');
				if (eqIdx < 0) continue;
				const key = opt.substring(0, eqIdx).trim().toLowerCase();
				const value = opt.substring(eqIdx + 1).trim().toLowerCase();
				if (key === 'mode') {
					if (value === 'tabbed' || value === 'stacked') {
						options.mode = value;
					}
				} else if (key === 'formatting') {
					options.formatting = value !== 'off';
				}
			}
		}

		return { ranges, options };
	}

	private async renderSpreadsheetBlock(source: string, el: HTMLElement, ctx: MarkdownPostProcessorContext) {
		const match = source.trim().match(/^(.+?)\((.+?)\)$/);
		if (!match) {
			throw new Error('Invalid format. Expected: filename.xlsx(A1:B2) or filename.xlsx(A1:B2,C1:D5; mode=tabbed)');
		}

		const [_, filename, innerArg] = match;
		const spec = this.parseSpec(innerArg);

		// Resolve the file path. Obsidian mobile sandbox can't use Node's `path` module,
		// so resolve with string ops. `sourcePath` always uses forward slashes.
		const notePath = ctx.sourcePath;
		const lastSlash = notePath.lastIndexOf('/');
		const noteDir = lastSlash >= 0 ? notePath.substring(0, lastSlash) : '';
		const filePath = noteDir ? `${noteDir}/${filename}` : filename;

		const abstractFile = this.app.vault.getAbstractFileByPath(filePath);
		if (!abstractFile || !(abstractFile instanceof TFile)) {
			throw new Error(`File not found or is not a valid file: ${filename} (looking in ${filePath})`);
		}
		const file = abstractFile;

		const arrayBuffer = await this.app.vault.readBinary(file);
		const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });

		const handle: WorksheetHandle = { workbook, file, filePath, filename };

		// Outer block
		const block = el.createEl('div', { cls: 'spreadsheet-block' });
		// Body wrapper — contains all sub-blocks. Cleared and re-rendered by
		// the "Load all sheets" button so we can swap content without losing
		// the outer toolbar (and its Open-in-Excel handler).
		const body = block.createEl('div', { cls: 'spreadsheet-body' });

		// Top toolbar — filename + open-in-excel action (always shown).
		this.renderTopToolbar(block, handle, body);

		if (spec.options.mode === 'tabbed') {
			this.renderTabbedBlock(body, handle, spec.ranges, spec.options);
		} else {
			for (const range of spec.ranges) {
				this.renderSubBlock(body, handle, range, spec.options);
			}
		}
	}

	private renderTopToolbar(parent: HTMLElement, handle: WorksheetHandle, body: HTMLElement) {
		const toolbar = parent.createEl('div', { cls: 'spreadsheet-toolbar' });
		const label = toolbar.createEl('span', {
			cls: 'spreadsheet-label',
			text: handle.filename,
		});
		label.title = handle.filePath;

		const actions = toolbar.createEl('span', { cls: 'spreadsheet-actions' });

		// Load-all-sheets action — hidden when the workbook only has one sheet.
		if (handle.workbook.SheetNames.length > 1) {
			const loadAllLink = actions.createEl('a', {
				cls: 'spreadsheet-load-all',
				text: 'Load all sheets',
				href: '#',
			});
			loadAllLink.addEventListener('click', async (evt) => {
				evt.preventDefault();
				await this.loadAllSheets(body, handle);
			});
		}

		const openLink = actions.createEl('a', {
			cls: 'spreadsheet-open',
			text: 'Open in Excel',
			href: '#',
		});
		openLink.addEventListener('click', async (evt) => {
			evt.preventDefault();
			const app = this.app as any;
			try {
				if (typeof app.openWithDefaultApp === 'function') {
					await app.openWithDefaultApp(handle.file.path);
				} else if (typeof app.showInFolder === 'function') {
					app.showInFolder(handle.file.path);
				} else if (typeof app.revealInFolder === 'function') {
					app.revealInFolder(handle.file);
				} else {
					throw new Error('No file-open API available in this Obsidian version');
				}
			} catch (err) {
				console.error('SpreadsheetSync: openWithDefaultApp failed', err);
				const errDiv = parent.createEl('div', {
					text: `Could not open ${handle.filename}: ${err.message}. File path: ${handle.filePath}`,
					cls: 'spreadsheet-error',
				});
				setTimeout(() => errDiv.remove(), 8000);
			}
		});
	}

	private async loadAllSheets(body: HTMLElement, handle: WorksheetHandle) {
		// Re-read the file so newly-added sheets show up.
		const arrayBuffer = await this.app.vault.readBinary(handle.file);
		const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
		handle.workbook = workbook;

		body.empty();
		const options: RenderOptions = { mode: 'stacked', formatting: true };
		let rendered = 0;
		for (const sheetName of workbook.SheetNames) {
			const worksheet = workbook.Sheets[sheetName];
			if (!worksheet || !worksheet['!ref']) continue; // skip empty / hidden
			const ref = XLSX.utils.decode_range(worksheet['!ref']);
			if (ref.e.r < ref.s.r || ref.e.c < ref.s.c) continue; // skip empty ranges
			const range: SpreadsheetRange = {
				sheet: sheetName,
				startCell: XLSX.utils.encode_cell({ r: ref.s.r, c: ref.s.c }),
				endCell: XLSX.utils.encode_cell({ r: ref.e.r, c: ref.e.c }),
			};
			this.renderSubBlock(body, handle, range, options);
			rendered++;
		}
		if (rendered === 0) {
			body.createEl('div', {
				text: 'No sheets with data found in this workbook.',
				cls: 'spreadsheet-error',
			});
		}
	}

	private renderSubBlock(
		parent: HTMLElement,
		handle: WorksheetHandle,
		range: SpreadsheetRange,
		options: RenderOptions,
	) {
		const sub = parent.createEl('div', { cls: 'spreadsheet-subblock' });

		// Per-range toolbar — sheet + range label.
		const subBar = sub.createEl('div', { cls: 'spreadsheet-subtoolbar' });
		const sheetName = range.sheet || handle.workbook.SheetNames[0];
		subBar.createEl('span', {
			cls: 'spreadsheet-label',
			text: `${sheetName} · ${range.startCell}:${range.endCell}`,
		});

		// Scroll wrapper.
		const scroll = sub.createEl('div', { cls: 'spreadsheet-scroll' });

		// Footer with Refresh + Use full range.
		const footer = sub.createEl('div', { cls: 'spreadsheet-footer' });
		this.renderFooter(footer, sub, handle, range, options);

		// Table itself.
		const table = scroll.createEl('table', { cls: 'spreadsheet-table' });
		const worksheet = handle.workbook.Sheets[sheetName];
		if (!worksheet) {
			sub.createEl('div', { text: `Sheet not found: ${sheetName}`, cls: 'spreadsheet-error' });
			return;
		}
		this.renderTable(table, worksheet, range, options.formatting);
	}

	private renderTabbedBlock(
		parent: HTMLElement,
		handle: WorksheetHandle,
		ranges: SpreadsheetRange[],
		options: RenderOptions,
	) {
		const tabBar = parent.createEl('div', { cls: 'spreadsheet-tabbar' });
		const subHosts: HTMLElement[] = [];
		const tabs: HTMLElement[] = [];

		ranges.forEach((range, idx) => {
			const sheetName = range.sheet || handle.workbook.SheetNames[0];
			const tabLabel = `${sheetName} · ${range.startCell}:${range.endCell}`;

			const tab = tabBar.createEl('button', {
				cls: 'spreadsheet-tab',
				text: tabLabel,
				attr: { type: 'button' },
			});
			if (idx === 0) tab.addClass('spreadsheet-tab-active');
			tabs.push(tab);

			const sub = parent.createEl('div', { cls: 'spreadsheet-subblock spreadsheet-tab-panel' });
			if (idx > 0) sub.style.display = 'none';
			subHosts.push(sub);

			tab.addEventListener('click', () => {
				subHosts.forEach((s, i) => {
					s.style.display = i === idx ? '' : 'none';
					tabs[i].removeClass('spreadsheet-tab-active');
				});
				tab.addClass('spreadsheet-tab-active');
			});

			// Sub-toolbar
			const subBar = sub.createEl('div', { cls: 'spreadsheet-subtoolbar' });
			subBar.createEl('span', { cls: 'spreadsheet-label', text: tabLabel });

			// Scroll wrapper
			const scroll = sub.createEl('div', { cls: 'spreadsheet-scroll' });

			// Footer
			const footer = sub.createEl('div', { cls: 'spreadsheet-footer' });
			this.renderFooter(footer, sub, handle, range, options);

			// Table
			const table = scroll.createEl('table', { cls: 'spreadsheet-table' });
			const worksheet = handle.workbook.Sheets[sheetName];
			if (!worksheet) {
				sub.createEl('div', { text: `Sheet not found: ${sheetName}`, cls: 'spreadsheet-error' });
				return;
			}
			this.renderTable(table, worksheet, range, options.formatting);
		});
	}

	private renderFooter(
		footer: HTMLElement,
		sub: HTMLElement,
		handle: WorksheetHandle,
		range: SpreadsheetRange,
		options: RenderOptions,
	) {
		footer.createEl('span', {
			cls: 'spreadsheet-label',
			text: `Range ${range.startCell}:${range.endCell}`,
		});

		const refreshBtn = footer.createEl('a', { text: 'Refresh', href: '#' });
		refreshBtn.addEventListener('click', async (evt) => {
			evt.preventDefault();
			await this.refreshSubBlock(sub, handle, range, options);
		});

		const fullBtn = footer.createEl('a', { text: 'Use full range', href: '#' });
		fullBtn.addEventListener('click', async (evt) => {
			evt.preventDefault();
			const sheetName = range.sheet || handle.workbook.SheetNames[0];
			const worksheet = handle.workbook.Sheets[sheetName];
			if (!worksheet || !worksheet['!ref']) return;
			const ref = XLSX.utils.decode_range(worksheet['!ref']);
			const fullRange: SpreadsheetRange = {
				sheet: range.sheet,
				startCell: XLSX.utils.encode_cell({ r: ref.s.r, c: ref.s.c }),
				endCell: XLSX.utils.encode_cell({ r: ref.e.r, c: ref.e.c }),
			};
			await this.refreshSubBlock(sub, handle, fullRange, options);
		});
	}

	private async refreshSubBlock(
		sub: HTMLElement,
		handle: WorksheetHandle,
		range: SpreadsheetRange,
		options: RenderOptions,
	) {
		// Re-read the file (may have changed on disk) and replace just the table.
		const arrayBuffer = await this.app.vault.readBinary(handle.file);
		const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
		handle.workbook = workbook;

		const sheetName = range.sheet || workbook.SheetNames[0];
		const worksheet = workbook.Sheets[sheetName];
		if (!worksheet) {
			sub.createEl('div', { text: `Sheet not found: ${sheetName}`, cls: 'spreadsheet-error' });
			return;
		}

		// Update the sub-toolbar label.
		const subBar = sub.querySelector('.spreadsheet-subtoolbar .spreadsheet-label');
		if (subBar) subBar.textContent = `${sheetName} · ${range.startCell}:${range.endCell}`;

		// Update the footer label.
		const footerLabel = sub.querySelector('.spreadsheet-footer .spreadsheet-label');
		if (footerLabel) footerLabel.textContent = `Range ${range.startCell}:${range.endCell}`;

		// Replace the table.
		const oldScroll = sub.querySelector('.spreadsheet-scroll');
		if (oldScroll) oldScroll.empty();
		const scroll = oldScroll || sub.createEl('div', { cls: 'spreadsheet-scroll' });
		const table = scroll.createEl('table', { cls: 'spreadsheet-table' });
		this.renderTable(table, worksheet, range, options.formatting);
	}

	private renderTable(
		table: HTMLElement,
		worksheet: XLSX.WorkSheet,
		range: SpreadsheetRange,
		formattingOn: boolean,
	) {
		const rangeInfo = XLSX.utils.decode_range(`${range.startCell}:${range.endCell}`);
		const headerRow = rangeInfo.s.r;

		for (let r = rangeInfo.s.r; r <= rangeInfo.e.r; r++) {
			const tr = table.createEl('tr');
			const isHeader = r === headerRow;
			for (let c = rangeInfo.s.c; c <= rangeInfo.e.c; c++) {
				const cellRef = XLSX.utils.encode_cell({ r, c });
				const cell = worksheet[cellRef];
				const td = tr.createEl(isHeader ? 'th' : 'td');
				if (cell?.v != null) {
					td.textContent = String(cell.v);
				}
				if (formattingOn && cell?.s) {
					this.applyCellStyle(td, cell.s);
				}
			}
		}
	}

	private applyCellStyle(td: HTMLElement, style: any) {
		const fillRgb = style?.fill?.fgColor?.rgb;
		if (fillRgb && typeof fillRgb === 'string') {
			// SheetJS may give 8-char ARGB; CSS wants 6-char RGB with a leading '#'.
			td.style.backgroundColor = '#' + fillRgb.slice(-6);
		}
		const fontRgb = style?.font?.color?.rgb;
		if (fontRgb && typeof fontRgb === 'string') {
			td.style.color = '#' + fontRgb.slice(-6);
		}
		if (style?.font?.bold) td.style.fontWeight = 'bold';
		if (style?.font?.italic) td.style.fontStyle = 'italic';
	}
}