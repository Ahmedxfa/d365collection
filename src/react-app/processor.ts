import { unzipSync } from "fflate";
import Papa from "papaparse";

export interface PeriodCustomer {
	account: string;
	name: string;
	customerGroup: string;
	buckets: number[];
	total: number;
}

export interface PeriodData {
	bucketNames: string[];
	customers: Map<string, PeriodCustomer>;
	reportDate: string;
}

export interface MergedCustomer {
	account: string;
	name: string;
	customerGroup: string;
	segment: string;
	buckets: number[];
	total: number;
}

export interface ProcessingResult {
	bucketNames: string[];
	rows: MergedCustomer[];
	total: number;
	segmentMatches: number;
	reportDate: string;
}

function normalize(value: string): string {
	return value.trim().toLocaleLowerCase();
}

function accountKey(value: string): string {
	return value.trim().toLocaleUpperCase();
}

function isAccountId(value: string): boolean {
	return /^cst[\da-z]+$/i.test(value.trim());
}

function parseAmount(value: string | undefined, fileName: string, rowNumber: number): number {
	const cleaned = (value ?? "")
		.trim()
		.replace(/^\((.*)\)$/, "-$1")
		.replace(/[,\s\u00a0]/g, "");
	if (cleaned === "" || cleaned === "-") return 0;
	const amount = Number(cleaned);
	if (!Number.isFinite(amount)) {
		throw new Error(`قيمة رصيد غير صالحة في ${fileName}، الصف ${rowNumber}.`);
	}
	return amount;
}

function parseReportDate(value: string | undefined, fileName: string): string {
	const match = value?.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
	if (!match) {
		throw new Error(`تعذر تحديد تاريخ الرصيد من حقل Balance as of في ${fileName}.`);
	}
	const [, dayText, monthText, yearText] = match;
	const day = Number(dayText);
	const month = Number(monthText);
	const year = Number(yearText);
	const date = new Date(Date.UTC(year, month - 1, day));
	if (
		date.getUTCFullYear() !== year ||
		date.getUTCMonth() !== month - 1 ||
		date.getUTCDate() !== day
	) {
		throw new Error(`تاريخ الرصيد غير صالح في ${fileName}.`);
	}
	return `${yearText}-${monthText.padStart(2, "0")}-${dayText.padStart(2, "0")}`;
}

async function readPeriod(file: File, verifyReportTotal: boolean): Promise<PeriodData> {
	const parsed = Papa.parse<string[]>(await file.text(), {
		skipEmptyLines: "greedy",
	});
	if (parsed.errors.length > 0) {
		throw new Error(`تعذرت قراءة ملف الفترة ${file.name}: ${parsed.errors[0].message}`);
	}

	const markerRow = parsed.data.find((row) =>
		row.some((cell) => normalize(cell) === "account"),
	);
	if (!markerRow) {
		throw new Error(`لم أجد عمود Account في ملف الفترة ${file.name}.`);
	}

	const markerIndex = markerRow.findIndex((cell) => normalize(cell) === "account");
	const accountIndex = markerRow.findIndex(
		(cell, index) => index > markerIndex && isAccountId(cell),
	);
	if (accountIndex < 0 || accountIndex + 9 >= markerRow.length) {
		throw new Error(`تنسيق ملف الفترة ${file.name} غير متوقع؛ تعذر تحديد أعمدة الحسابات والأرصدة.`);
	}
	const balanceDateLabelIndex = markerRow.findIndex((cell) => normalize(cell) === "balance as of");
	if (balanceDateLabelIndex < 0) {
		throw new Error(`لم أجد حقل Balance as of في ملف الفترة ${file.name}.`);
	}
	const reportDate = parseReportDate(markerRow[balanceDateLabelIndex + 1], file.name);

	const bucketNames = markerRow
		.slice(0, 6)
		.map((label, index) => label.trim() || `الفئة ${index + 1}`);
	const customers = new Map<string, PeriodCustomer>();

	for (const [index, row] of parsed.data.entries()) {
		const rawAccount = row[accountIndex]?.trim() ?? "";
		if (!isAccountId(rawAccount)) continue;

		const key = accountKey(rawAccount);
		if (customers.has(key)) {
			throw new Error(`الحساب ${rawAccount} مكرر في ملف الفترة ${file.name}.`);
		}
		const buckets = Array.from({ length: 6 }, (_, bucketIndex) =>
			parseAmount(row[accountIndex + 4 + bucketIndex], file.name, index + 1),
		);
		customers.set(key, {
			account: rawAccount,
			name: row[accountIndex + 1]?.trim() ?? "",
			customerGroup: row[accountIndex + 2]?.trim() ?? "",
			buckets,
			total: verifyReportTotal
				? parseAmount(row[accountIndex + 3], file.name, index + 1)
				: 0,
		});
	}

	if (customers.size === 0) {
		throw new Error(`لم أجد أي حسابات عملاء في ملف الفترة ${file.name}.`);
	}
	if (verifyReportTotal) {
		const customerTotal = [...customers.values()].reduce((sum, customer) => sum + customer.total, 0);
		const reportTotal = parseAmount(markerRow[accountIndex + 11], file.name, 1);
		if (Math.abs(customerTotal - reportTotal) > 0.02) {
			throw new Error(
				`مجموع أرصدة الحسابات في ${file.name} لا يطابق الإجمالي المطبوع في التقرير (${reportTotal.toFixed(2)}).`,
			);
		}
	}

	return { bucketNames, customers, reportDate };
}

const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

function parseXml(source: string, fileName: string): XMLDocument {
	const document = new DOMParser().parseFromString(source, "application/xml");
	if (document.getElementsByTagName("parsererror").length > 0) {
		throw new Error(`ملف Excel غير صالح: ${fileName}.`);
	}
	return document;
}

function normalizeZipPath(base: string, target: string): string {
	const parts = (target.startsWith("/") ? [] : base.split("/")).concat(target.split("/"));
	const normalized: string[] = [];
	for (const part of parts) {
		if (part === "" || part === ".") continue;
		if (part === "..") normalized.pop();
		else normalized.push(part);
	}
	return normalized.join("/");
}

function columnNumber(reference: string): number {
	const letters = reference.match(/^[A-Z]+/i)?.[0]?.toUpperCase();
	if (!letters) return 0;
	return [...letters].reduce(
		(value, letter) => value * 26 + letter.charCodeAt(0) - 64,
		0,
	);
}

function xmlText(element: Element): string {
	return Array.from(element.getElementsByTagNameNS(SPREADSHEET_NS, "t"))
		.map((text) => text.textContent ?? "")
		.join("");
}

function xlsxRows(source: string, sharedStrings: string[], fileName: string): Map<number, Map<number, string>> {
	const document = parseXml(source, fileName);
	const rows = new Map<number, Map<number, string>>();
	const rowElements = document.getElementsByTagNameNS(SPREADSHEET_NS, "row");
	for (const rowElement of Array.from(rowElements)) {
		const rowNumber = Number(rowElement.getAttribute("r"));
		const cells = new Map<number, string>();
		let nextColumn = 1;
		for (const cell of Array.from(rowElement.getElementsByTagNameNS(SPREADSHEET_NS, "c"))) {
			const column = columnNumber(cell.getAttribute("r") ?? "") || nextColumn;
			nextColumn = column + 1;
			if (column === 0) continue;
			const type = cell.getAttribute("t");
			if (type === "inlineStr") {
				cells.set(column, xmlText(cell).trim());
				continue;
			}
			const value = cell.getElementsByTagNameNS(SPREADSHEET_NS, "v").item(0)?.textContent ?? "";
			if (type === "s" && value !== "") {
				const sharedString = sharedStrings[Number(value)];
				if (sharedString === undefined) {
					throw new Error(`مرجع نص غير صالح في ملف السيجمينت ${fileName}.`);
				}
				cells.set(column, sharedString.trim());
			} else {
				cells.set(column, value.trim());
			}
		}
		if (rowNumber > 0) rows.set(rowNumber, cells);
	}
	return rows;
}

async function readSegments(file: File): Promise<Map<string, string>> {
	const archive = unzipSync(new Uint8Array(await file.arrayBuffer()));
	const getFile = (path: string): string => {
		const entry = Object.prototype.hasOwnProperty.call(archive, path) ? archive[path] : undefined;
		if (!entry) throw new Error(`ملف Excel ناقص أو غير صالح: ${file.name}.`);
		return new TextDecoder().decode(entry);
	};
	const workbook = parseXml(getFile("xl/workbook.xml"), file.name);
	const relations = parseXml(getFile("xl/_rels/workbook.xml.rels"), file.name);
	const relationElements = Array.from(relations.getElementsByTagNameNS(PACKAGE_REL_NS, "Relationship"));
	const sheets = Array.from(workbook.getElementsByTagNameNS(SPREADSHEET_NS, "sheet"));
	const visibleSheet = sheets.find(
		(sheet) => sheet.getAttribute("state") !== "hidden" && sheet.getAttribute("state") !== "veryHidden",
	);
	if (!visibleSheet) {
		throw new Error("ملف السيجمينت لا يحتوي على ورقة عمل ظاهرة.");
	}

	const relationId = visibleSheet.getAttributeNS(OFFICE_REL_NS, "id") ?? visibleSheet.getAttribute("r:id");
	const relationship = relationElements.find((item) => item.getAttribute("Id") === relationId);
	const target = relationship?.getAttribute("Target");
	if (!target) throw new Error(`تعذر تحديد ورقة العمل في ملف السيجمينت ${file.name}.`);
	const worksheetPath = normalizeZipPath("xl", target);
	const sharedStringsPath = "xl/sharedStrings.xml";
	const sharedStrings = Object.prototype.hasOwnProperty.call(archive, sharedStringsPath)
		? Array.from(
				parseXml(getFile(sharedStringsPath), file.name).getElementsByTagNameNS(SPREADSHEET_NS, "si"),
			).map(xmlText)
		: [];
	const worksheetRows = xlsxRows(getFile(worksheetPath), sharedStrings, file.name);

	let headerRow = 0;
	let accountColumn = 0;
	let segmentColumn = 0;
	for (let rowNumber = 1; rowNumber <= 15; rowNumber += 1) {
		const row = worksheetRows.get(rowNumber);
		for (const [column, value] of row ?? []) {
			const header = normalize(value).replace(/[^a-z]/g, "");
			if (header === "customeraccount") accountColumn = column;
			if (header === "segment") segmentColumn = column;
		}
		if (accountColumn > 0 && segmentColumn > 0) {
			headerRow = rowNumber;
			break;
		}
	}
	if (!headerRow) {
		throw new Error("لم أجد عمودي Customer account وSegment في ملف السيجمينت.");
	}

	const segments = new Map<string, string>();
	for (const [rowNumber, row] of worksheetRows) {
		if (rowNumber <= headerRow) continue;
		const account = row.get(accountColumn) ?? "";
		if (!isAccountId(account)) continue;

		const key = accountKey(account);
		if (segments.has(key)) {
			throw new Error(`الحساب ${account} مكرر في ملف السيجمينت.`);
		}
		segments.set(key, row.get(segmentColumn) ?? "");
	}
	if (segments.size === 0) {
		throw new Error("لم أجد حسابات في ملف السيجمينت.");
	}
	return segments;
}

export async function processFiles(
	periodOneFile: File,
	periodTwoFile: File,
	segmentFile: File,
): Promise<ProcessingResult> {
	const [periodOne, periodTwo, segments] = await Promise.all([
		readPeriod(periodOneFile, false),
		readPeriod(periodTwoFile, true),
		readSegments(segmentFile),
	]);
	if (periodOne.reportDate !== periodTwo.reportDate) {
		throw new Error(
			`تاريخ الرصيد مختلف بين الملفين: ${periodOne.reportDate} في الملف الأول و${periodTwo.reportDate} في الملف الثاني. يجب أن يكونا لنفس التاريخ لدمج الفترات.`,
		);
	}

	const periodOneOnly = [...periodOne.customers.keys()].filter((key) => !periodTwo.customers.has(key));
	const periodTwoOnly = [...periodTwo.customers.keys()].filter((key) => !periodOne.customers.has(key));
	if (periodOneOnly.length > 0 || periodTwoOnly.length > 0) {
		throw new Error(
			`قوائم الحسابات لا تتطابق بين الملفين: ${periodOneOnly.length} حساب في الأول فقط و${periodTwoOnly.length} في الثاني فقط.`,
		);
	}

	const rows = [...periodOne.customers.entries()].map(([key, first]): MergedCustomer => {
		const second = periodTwo.customers.get(key);
		if (!second) throw new Error(`تعذر العثور على الحساب ${first.account} في ملف الفترات الثاني.`);

		const buckets = [...first.buckets, ...second.buckets.slice(3)];
		const total = second.total;
		const bucketTotal = buckets.reduce((sum, amount) => sum + amount, 0);
		if (Math.abs(bucketTotal - total) > 0.02) {
			throw new Error(`مجموع الفترات لا يطابق الرصيد الفعلي للحساب ${first.account}.`);
		}

		const segment = segments.get(key) ?? "";
		return {
			account: first.account,
			name: first.name || second.name,
			customerGroup: first.customerGroup || second.customerGroup,
			segment: segment || "غير محدد",
			buckets,
			total,
		};
	});
	const total = rows.reduce((sum, row) => sum + row.total, 0);

	return {
		bucketNames: [...periodOne.bucketNames, ...periodTwo.bucketNames.slice(3)],
		rows,
		total,
		segmentMatches: rows.filter((row) => Boolean(segments.get(accountKey(row.account)))).length,
		reportDate: periodOne.reportDate,
	};
}
