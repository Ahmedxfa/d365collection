import { Hono } from "hono";

interface Bindings extends Env {
	DB: D1Database;
}

interface CustomerRow {
	account: string;
	name: string;
	customerGroup: string;
	segment: string;
	buckets: number[];
	total: number;
}

interface ArchivePayload {
	bucketNames: string[];
	rows: CustomerRow[];
	total: number;
	segmentMatches: number;
}

interface ArchiveSummary {
	id: string;
	createdAt: string;
	total: number;
	rowCount: number;
	segmentMatches: number;
	bucketNames: string[];
}

interface ArchiveRecord {
	id: string;
	created_at: string;
	total: number;
	row_count: number;
	segment_matches: number;
	bucket_names_json: string;
}

const MAX_BODY_LENGTH = 5_000_000;
const MAX_ROWS = 5_000;
const MAX_BUCKETS = 20;
const ROW_BATCH_SIZE = 100;

const app = new Hono<{ Bindings: Bindings }>();

function jsonError(message: string, status: 400 | 404 | 413 | 500) {
	return Response.json({ error: message }, { status });
}

function validString(value: unknown, maxLength = 300): value is string {
	return typeof value === "string" && value.length <= maxLength;
}

function finiteNumber(value: unknown): value is number {
	return typeof value === "number" && Number.isFinite(value);
}

function validatePayload(value: unknown): ArchivePayload | null {
	if (!value || typeof value !== "object") return null;
	const payload = value as Partial<ArchivePayload>;
	if (
		!Array.isArray(payload.bucketNames) ||
		payload.bucketNames.length === 0 ||
		payload.bucketNames.length > MAX_BUCKETS ||
		!payload.bucketNames.every((name) => validString(name, 100)) ||
		!Array.isArray(payload.rows) ||
		payload.rows.length === 0 ||
		payload.rows.length > MAX_ROWS ||
		!finiteNumber(payload.total) ||
		typeof payload.segmentMatches !== "number" ||
		!Number.isInteger(payload.segmentMatches) ||
		payload.segmentMatches < 0 ||
		payload.segmentMatches > payload.rows.length
	) {
		return null;
	}
	const accounts = new Set<string>();
	for (const row of payload.rows) {
		if (
			!row ||
			!validString(row.account, 100) ||
			!validString(row.name) ||
			!validString(row.customerGroup) ||
			!validString(row.segment) ||
			!Array.isArray(row.buckets) ||
			row.buckets.length !== payload.bucketNames.length ||
			!row.buckets.every(finiteNumber) ||
			!finiteNumber(row.total) ||
			accounts.has(row.account)
		) {
			return null;
		}
		accounts.add(row.account);
	}
	return payload as ArchivePayload;
}

function toSummary(record: ArchiveRecord): ArchiveSummary {
	return {
		id: record.id,
		createdAt: record.created_at,
		total: record.total,
		rowCount: record.row_count,
		segmentMatches: record.segment_matches,
		bucketNames: JSON.parse(record.bucket_names_json) as string[],
	};
}

function isValidArchiveId(id: string): boolean {
	return /^[0-9a-f-]{36}$/i.test(id);
}

app.get("/api/archives", async (context) => {
	try {
		const { results } = await context.env.DB.prepare(
			"SELECT id, created_at, total, row_count, segment_matches, bucket_names_json FROM archives ORDER BY created_at DESC",
		).all<ArchiveRecord>();
		return context.json({ archives: results.map(toSummary) });
	} catch (error) {
		console.error("Failed to list archives", error);
		return jsonError("تعذر تحميل أرشيف البيانات.", 500);
	}
});

app.get("/api/archives/:id", async (context) => {
	const { id } = context.req.param();
	if (!isValidArchiveId(id)) return jsonError("معرّف النسخة غير صالح.", 400);
	try {
		const record = await context.env.DB.prepare(
			"SELECT id, created_at, total, row_count, segment_matches, bucket_names_json FROM archives WHERE id = ?",
		).bind(id).first<ArchiveRecord>();
		if (!record) return jsonError("لم يتم العثور على النسخة المطلوبة.", 404);
		const { results } = await context.env.DB.prepare(
			"SELECT account, name, customer_group, segment, buckets_json, total FROM archive_rows WHERE archive_id = ? ORDER BY account",
		).bind(id).all<{
			account: string;
			name: string;
			customer_group: string;
			segment: string;
			buckets_json: string;
			total: number;
		}>();
		return context.json({
			...toSummary(record),
			rows: results.map((row) => ({
				account: row.account,
				name: row.name,
				customerGroup: row.customer_group,
				segment: row.segment,
				buckets: JSON.parse(row.buckets_json) as number[],
				total: row.total,
			})),
		});
	} catch (error) {
		console.error("Failed to load archive", error);
		return jsonError("تعذر تحميل نسخة البيانات.", 500);
	}
});

app.post("/api/archives", async (context) => {
	const contentLength = Number(context.req.header("content-length") ?? 0);
	if (contentLength > MAX_BODY_LENGTH) return jsonError("حجم الملف أكبر من الحد المسموح.", 413);
	const bodyText = await context.req.text();
	if (bodyText.length > MAX_BODY_LENGTH) return jsonError("حجم الملف أكبر من الحد المسموح.", 413);
	let body: unknown;
	try {
		body = JSON.parse(bodyText);
	} catch {
		return jsonError("بيانات النسخة غير صالحة.", 400);
	}
	const payload = validatePayload(body);
	if (!payload) return jsonError("بيانات النسخة غير مكتملة أو غير صالحة.", 400);

	const id = crypto.randomUUID();
	const createdAt = new Date().toISOString();
	const total = payload.rows.reduce((sum, row) => sum + row.total, 0);
	if (Math.abs(total - payload.total) > 0.02) {
		return jsonError("إجمالي النسخة لا يطابق مجموع أرصدة الحسابات.", 400);
	}

	try {
		await context.env.DB.prepare(
			"INSERT INTO archives (id, created_at, total, row_count, segment_matches, bucket_names_json) VALUES (?, ?, ?, ?, ?, ?)",
		).bind(id, createdAt, total, payload.rows.length, payload.segmentMatches, JSON.stringify(payload.bucketNames)).run();
		for (let start = 0; start < payload.rows.length; start += ROW_BATCH_SIZE) {
			const statements = payload.rows.slice(start, start + ROW_BATCH_SIZE).map((row) =>
				context.env.DB.prepare(
					"INSERT INTO archive_rows (archive_id, account, name, customer_group, segment, buckets_json, total) VALUES (?, ?, ?, ?, ?, ?, ?)",
				).bind(id, row.account, row.name, row.customerGroup, row.segment, JSON.stringify(row.buckets), row.total),
			);
			await context.env.DB.batch(statements);
		}
		return context.json({
			archive: { id, createdAt, total, rowCount: payload.rows.length, segmentMatches: payload.segmentMatches, bucketNames: payload.bucketNames },
		}, 201);
	} catch (error) {
		try {
			await context.env.DB.prepare("DELETE FROM archives WHERE id = ?").bind(id).run();
		} catch (cleanupError) {
			console.error("Failed to clean up incomplete archive", cleanupError);
		}
		console.error("Failed to save archive", error);
		return jsonError("تعذر حفظ النسخة في الأرشيف.", 500);
	}
});

app.delete("/api/archives/:id", async (context) => {
	const { id } = context.req.param();
	if (!isValidArchiveId(id)) return jsonError("معرّف النسخة غير صالح.", 400);
	try {
		const result = await context.env.DB.prepare("DELETE FROM archives WHERE id = ?").bind(id).run();
		if (!result.meta.changes) return jsonError("لم يتم العثور على النسخة المطلوبة.", 404);
		await context.env.DB.prepare("DELETE FROM archive_rows WHERE archive_id = ?").bind(id).run();
		return context.body(null, 204);
	} catch (error) {
		console.error("Failed to delete archive", error);
		return jsonError("تعذر حذف النسخة.", 500);
	}
});

app.get("/api/", (context) => context.json({ name: "Cloudflare" }));

export default app;
