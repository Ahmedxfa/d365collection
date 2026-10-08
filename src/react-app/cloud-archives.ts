import type { ProcessingResult } from "./processor";

export interface ArchiveSummary {
	id: string;
	createdAt: string;
	total: number;
	rowCount: number;
	segmentMatches: number;
	bucketNames: string[];
}

export interface ArchivedDataset extends ArchiveSummary {
	rows: ProcessingResult["rows"];
}

async function readResponse<T>(response: Response): Promise<T> {
	if (!response.ok) {
		let message = `تعذر تنفيذ طلب الأرشيف (${response.status}).`;
		try {
			const body = await response.json() as { error?: unknown };
			if (typeof body.error === "string") message = body.error;
		} catch {
			// Use the explicit HTTP status when the response is not JSON.
		}
		throw new Error(message);
	}
	return response.json() as Promise<T>;
}

export async function listArchives(): Promise<ArchiveSummary[]> {
	const response = await fetch("/api/archives", { headers: { Accept: "application/json" } });
	const data = await readResponse<{ archives: ArchiveSummary[] }>(response);
	return data.archives;
}

export async function loadArchive(id: string): Promise<ArchivedDataset> {
	const response = await fetch(`/api/archives/${encodeURIComponent(id)}`, {
		headers: { Accept: "application/json" },
	});
	return readResponse<ArchivedDataset>(response);
}

export async function saveArchive(result: ProcessingResult): Promise<ArchiveSummary> {
	const response = await fetch("/api/archives", {
		method: "POST",
		headers: { "Content-Type": "application/json", Accept: "application/json" },
		body: JSON.stringify({
			bucketNames: result.bucketNames,
			rows: result.rows,
			total: result.total,
			segmentMatches: result.segmentMatches,
		}),
	});
	const data = await readResponse<{ archive: ArchiveSummary }>(response);
	return data.archive;
}

export async function deleteArchive(id: string): Promise<void> {
	const response = await fetch(`/api/archives/${encodeURIComponent(id)}`, { method: "DELETE" });
	if (!response.ok) {
		await readResponse<never>(response);
	}
}
