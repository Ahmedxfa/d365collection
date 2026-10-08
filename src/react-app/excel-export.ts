import { strToU8, zipSync } from "fflate";
import type { ProcessingResult } from "./processor";

const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function escapeXml(value: string): string {
	return value
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;")
		.replace(/'/g, "&apos;");
}

function columnName(index: number): string {
	let name = "";
	let number = index;
	while (number > 0) {
		number -= 1;
		name = String.fromCharCode(65 + (number % 26)) + name;
		number = Math.floor(number / 26);
	}
	return name;
}

function cellXml(reference: string, value: string | number, style: number): string {
	if (typeof value === "number") {
		return `<c r="${reference}" s="${style}" t="n"><v>${Number.isInteger(value) ? value : value.toFixed(2)}</v></c>`;
	}
	return `<c r="${reference}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function workbookFiles(result: ProcessingResult): Record<string, Uint8Array> {
	const headers = [
		"Customer account",
		"Customer name",
		"Customer group",
		"Segment",
		...result.bucketNames,
		"Total balance",
	];
	const rows = [
		headers,
		...result.rows.map((customer) => [
			customer.account,
			customer.name,
			customer.customerGroup,
			customer.segment,
			...customer.buckets,
			customer.total,
		]),
	];
	const lastColumn = columnName(headers.length);
	const columns = [
		{ min: 1, max: 1, width: 18 },
		{ min: 2, max: 2, width: 38 },
		{ min: 3, max: 3, width: 20 },
		{ min: 4, max: 4, width: 27 },
		{ min: 5, max: headers.length - 1, width: 18 },
		{ min: headers.length, max: headers.length, width: 20 },
	]
		.filter((column) => column.min <= column.max)
		.map((column) => `<col min="${column.min}" max="${column.max}" width="${column.width}" customWidth="1"/>`)
		.join("");
	const worksheetRows = rows
		.map(
			(row, rowIndex) =>
				`<row r="${rowIndex + 1}"${rowIndex === 0 ? ' ht="30" customHeight="1"' : ""}>${row
					.map((value, columnIndex) => {
						const style = rowIndex === 0 ? 1 : columnIndex >= 4 ? (columnIndex === headers.length - 1 ? 3 : 2) : 0;
						return cellXml(`${columnName(columnIndex + 1)}${rowIndex + 1}`, value, style);
					})
					.join("")}</row>`,
		)
		.join("");

	const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;
	const rootRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
	const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Merged aging periods" sheetId="1" r:id="rId1"/></sheets></workbook>`;
	const workbookRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
	const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/></font><font><b/><color rgb="FF165F4E"/><sz val="11"/><name val="Aptos"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF176F5E"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE5F2E9"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment horizontal="right"/></xf><xf numFmtId="4" fontId="2" fillId="3" borderId="0" xfId="0" applyNumberFormat="1"><alignment horizontal="right"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
	const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/><cols>${columns}</cols><sheetData>${worksheetRows}</sheetData><autoFilter ref="A1:${lastColumn}${rows.length}"/><pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>`;

	return {
		"[Content_Types].xml": strToU8(contentTypes),
		"_rels/.rels": strToU8(rootRelationships),
		"xl/workbook.xml": strToU8(workbook),
		"xl/_rels/workbook.xml.rels": strToU8(workbookRelationships),
		"xl/styles.xml": strToU8(styles),
		"xl/worksheets/sheet1.xml": strToU8(worksheet),
	};
}

export function createExcelFile(result: ProcessingResult): Blob {
	return new Blob([zipSync(workbookFiles(result))], { type: XLSX_CONTENT_TYPE });
}
