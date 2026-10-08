import { useEffect, useMemo, useRef, useState, type ChangeEvent, type KeyboardEvent, type PointerEvent as ReactPointerEvent } from "react";
import { createExcelFile } from "./excel-export";
import { processFiles, type ProcessingResult } from "./processor";
import { loadSavedDashboard, saveDashboard } from "./dashboard-storage";
import { deleteArchive, listArchives, loadArchive, saveArchive, type ArchiveSummary, type ArchivedDataset } from "./cloud-archives";
import "./App.css";

type UploadKey = "periodOne" | "periodTwo" | "segments";

interface Uploads {
	periodOne: File | null;
	periodTwo: File | null;
	segments: File | null;
}

const EMPTY_UPLOADS: Uploads = {
	periodOne: null,
	periodTwo: null,
	segments: null,
};

const OWNER_FILTERS = [
	{ id: "ammar", name: "عمار", startsWith: ["عمار"], contains: [] },
	{ id: "amjad", name: "أمجد", startsWith: ["امجد"], contains: [] },
	{ id: "ahmed", name: "أحمد", startsWith: ["احمد"], contains: [] },
	{ id: "emad", name: "عماد", startsWith: ["عماد"], contains: [] },
	{ id: "nawaz", name: "نواز", startsWith: [], contains: ["نواز"] },
	{ id: "mohammed-saad", name: "محمد سعد", startsWith: ["محمد سعد"], contains: [] },
] as const;

const PAGE_SIZE = 20;
const MIN_COLUMN_WIDTH = 80;
const INITIAL_COLUMN_WIDTHS = [260, 145, 150, ...Array(9).fill(125), 145];
const OVER_30_BUCKET_START = 5;
const OVER_45_BUCKET_START = 7;
const COMMISSION_RATE = 0.01;
const amountFormat = new Intl.NumberFormat("en-US", {
	minimumFractionDigits: 2,
	maximumFractionDigits: 2,
});
const percentFormat = new Intl.NumberFormat("ar", {
	minimumFractionDigits: 0,
	maximumFractionDigits: 2,
});
const dateFormat = new Intl.DateTimeFormat("ar", { dateStyle: "medium", timeStyle: "short" });
type SortColumn = "total" | "bucket";
type SortDirection = "asc" | "desc";
interface TableSort {
	column: SortColumn;
	bucketIndex?: number;
	direction: SortDirection;
}

interface SegmentComparison {
	segment: string;
	firstBalance: number;
	firstOver30: number;
	firstOver45: number;
	firstCommission: number;
	secondBalance: number;
	secondOver30: number;
	secondOver45: number;
	secondCommission: number;
}

function normalizeArabicName(value: string): string {
	return value
		.normalize("NFD")
		.replace(/[\u064B-\u065F\u0670\u0640]/g, "")
		.replace(/[أإآٱ]/g, "ا")
		.replace(/ى/g, "ي")
		.replace(/\s+/g, " ")
		.trim();
}

function ownerForSegment(segment: string): string | null {
	const normalized = normalizeArabicName(segment);
	return OWNER_FILTERS.find((owner) =>
		owner.startsWith.some((prefix) => normalized.startsWith(prefix)) ||
		owner.contains.some((fragment) => normalized.replace(/\s/g, "").includes(fragment)),
	)?.id ?? null;
}

function formatAmount(amount: number): string {
	return amountFormat.format(amount);
}

function formatPercent(amount: number, total: number): string {
	return `${percentFormat.format(total === 0 ? 0 : (amount / total) * 100)}٪`;
}

function sumBucketsFrom(customer: ProcessingResult["rows"][number], startIndex: number): number {
	return customer.buckets.slice(startIndex).reduce((sum, amount) => sum + amount, 0);
}

function downloadExcel(result: ProcessingResult): void {
	const url = URL.createObjectURL(createExcelFile(result));
	const link = document.createElement("a");
	link.href = url;
	link.download = "merged-aging-periods.xlsx";
	document.body.appendChild(link);
	link.click();
	window.setTimeout(() => {
		link.remove();
		URL.revokeObjectURL(url);
	}, 1000);
}

function HeaderUpload({
	title,
	accept,
	file,
	disabled,
	onChange,
}: {
	title: string;
	accept: string;
	file: File | null;
	disabled: boolean;
	onChange: (event: ChangeEvent<HTMLInputElement>) => void;
}) {
	return (
		<label className={`header-upload${file ? " selected" : ""}`}>
			<span className="header-upload-title">{title}</span>
			<span className="header-upload-choice">{file ? "✓ تغيير" : "＋ اختيار"}</span>
			<span className="header-upload-name" title={file?.name}>{file?.name ?? "لم يتم اختيار ملف"}</span>
			<input type="file" accept={accept} disabled={disabled} onChange={onChange} />
		</label>
	);
}

function App() {
	const [uploads, setUploads] = useState<Uploads>(EMPTY_UPLOADS);
	const [result, setResult] = useState<ProcessingResult | null>(null);
	const [lastUpdatedAt, setLastUpdatedAt] = useState<string | null>(null);
	const [currentArchiveId, setCurrentArchiveId] = useState<string | null>(null);
	const [archives, setArchives] = useState<ArchiveSummary[]>([]);
	const [archiveStatus, setArchiveStatus] = useState<"checking" | "connected" | "local">("checking");
	const [selectedArchiveIds, setSelectedArchiveIds] = useState<[string, string]>(["", ""]);
	const [comparison, setComparison] = useState<{
		first: ArchivedDataset;
		second: ArchivedDataset;
		rows: SegmentComparison[];
	} | null>(null);
	const [isComparing, setIsComparing] = useState(false);
	const [isRestoring, setIsRestoring] = useState(true);
	const [isProcessing, setIsProcessing] = useState(false);
	const [error, setError] = useState("");
	const [search, setSearch] = useState("");
	const [segmentFilter, setSegmentFilter] = useState("all");
	const [selectedOwners, setSelectedOwners] = useState<Set<string>>(
		() => new Set(OWNER_FILTERS.map((owner) => owner.id)),
	);
	const [showOtherSegments, setShowOtherSegments] = useState(false);
	const [page, setPage] = useState(1);
	const [sort, setSort] = useState<TableSort>({ column: "total", direction: "desc" });
	const [columnWidths, setColumnWidths] = useState(INITIAL_COLUMN_WIDTHS);
	const resizeStart = useRef<{ columnIndex: number; pointerX: number; width: number } | null>(null);
	const runId = useRef(0);

	useEffect(() => {
		let isCurrent = true;
		void (async () => {
			try {
				const archiveList = await listArchives();
				if (!isCurrent) return;
				setArchives(archiveList);
				setArchiveStatus("connected");
				if (archiveList.length > 0) {
					const latest = await loadArchive(archiveList[0].id);
					if (!isCurrent) return;
					setResult(latest);
					setLastUpdatedAt(latest.createdAt);
					setCurrentArchiveId(latest.id);
					setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
					setShowOtherSegments(false);
					setSelectedArchiveIds([archiveList[1]?.id ?? "", archiveList[0].id]);
					return;
				}
				const saved = await loadSavedDashboard();
				if (isCurrent) {
					if (saved) {
						setResult(saved.result);
						setLastUpdatedAt(saved.updatedAt);
						setArchiveStatus("local");
					}
				}
			} catch (remoteError: unknown) {
				if (!isCurrent) return;
				setArchiveStatus("local");
				try {
					const saved = await loadSavedDashboard();
					if (isCurrent && saved) {
						setResult(saved.result);
						setLastUpdatedAt(saved.updatedAt);
					}
				} catch (storageError: unknown) {
					setError(`تعذر استعادة النسخة السحابية والمحلية: ${storageError instanceof Error ? storageError.message : String(remoteError)}`);
				}
				if (isCurrent) {
					setError(`تعذر الاتصال بأرشيف Cloudflare: ${remoteError instanceof Error ? remoteError.message : "خطأ غير معروف."}`);
				}
			} finally {
				if (isCurrent) setIsRestoring(false);
			}
		})();
		return () => {
			isCurrent = false;
		};
	}, []);

	const ownerCounts = useMemo(() => {
		const counts = new Map<string, number>();
		for (const owner of OWNER_FILTERS) counts.set(owner.id, 0);
		for (const customer of result?.rows ?? []) {
			const ownerId = ownerForSegment(customer.segment);
			if (ownerId) counts.set(ownerId, (counts.get(ownerId) ?? 0) + 1);
		}
		return counts;
	}, [result]);
	const otherSegmentCount = useMemo(
		() => (result?.rows ?? []).filter((customer) => ownerForSegment(customer.segment) === null).length,
		[result],
	);

	const selectedRows = useMemo(() => {
		if (!result) return [];
		return result.rows.filter((customer) => {
			const ownerId = ownerForSegment(customer.segment);
			return ownerId
				? selectedOwners.has(ownerId)
				: showOtherSegments;
		});
	}, [result, selectedOwners, showOtherSegments]);

	const segments = useMemo(() => {
		if (!result) return [];
		const totalsBySegment = new Map<string, number>();
		for (const customer of selectedRows) {
			totalsBySegment.set(
				customer.segment,
				(totalsBySegment.get(customer.segment) ?? 0) + customer.total,
			);
		}
		return [...totalsBySegment.entries()].sort((a, b) => b[1] - a[1]);
	}, [result, selectedRows]);

	const bucketTotals = useMemo(() => {
		if (!result) return [];
		return result.bucketNames.map((name, index) => ({
			name,
			amount: selectedRows.reduce((sum, customer) => sum + customer.buckets[index], 0),
		}));
	}, [result, selectedRows]);

	const visibleRows = useMemo(() => {
		if (!result) return [];
		const query = search.trim().toLocaleLowerCase();
		const filtered = selectedRows.filter((customer) => {
			const matchesSegment = segmentFilter === "all" || customer.segment === segmentFilter;
			const matchesSearch =
				query === "" ||
				[customer.account, customer.name, customer.customerGroup, customer.segment]
					.join(" ")
					.toLocaleLowerCase()
					.includes(query);
			return matchesSegment && matchesSearch;
		});
		return [...filtered].sort((first, second) => {
			const firstValue = sort.column === "total"
				? first.total
				: first.buckets[sort.bucketIndex ?? 0] ?? 0;
			const secondValue = sort.column === "total"
				? second.total
				: second.buckets[sort.bucketIndex ?? 0] ?? 0;
			const comparison = firstValue - secondValue;
			return sort.direction === "desc" ? -comparison : comparison;
		});
	}, [result, search, segmentFilter, selectedRows, sort]);

	const selectedTotal = useMemo(
		() => selectedRows.reduce((sum, customer) => sum + customer.total, 0),
		[selectedRows],
	);
	const selectedOver30 = useMemo(
		() => selectedRows.reduce((sum, customer) => sum + sumBucketsFrom(customer, OVER_30_BUCKET_START), 0),
		[selectedRows],
	);
	const selectedOver45 = useMemo(
		() => selectedRows.reduce((sum, customer) => sum + sumBucketsFrom(customer, OVER_45_BUCKET_START), 0),
		[selectedRows],
	);
	const commissionBySegment = useMemo(() => {
		const totals = new Map<string, { total: number; over30: number; over45: number; count: number }>();
		for (const customer of selectedRows) {
			const current = totals.get(customer.segment) ?? { total: 0, over30: 0, over45: 0, count: 0 };
			totals.set(customer.segment, {
				total: current.total + customer.total,
				over30: current.over30 + sumBucketsFrom(customer, OVER_30_BUCKET_START),
				over45: current.over45 + sumBucketsFrom(customer, OVER_45_BUCKET_START),
				count: current.count + 1,
			});
		}
		return [...totals.entries()].sort((first, second) => second[1].over45 - first[1].over45);
	}, [selectedRows]);

	const pageCount = Math.max(1, Math.ceil(visibleRows.length / PAGE_SIZE));
	const pageRows = visibleRows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);
	const maxBucket = Math.max(0, ...bucketTotals.map((bucket) => bucket.amount));
	const maxSegment = Math.max(0, ...segments.slice(0, 8).map(([, amount]) => amount));
	const otherSegments = segments.slice(8);
	const otherSegmentTotal = otherSegments.reduce((sum, [, amount]) => sum + amount, 0);
	const formattedLastUpdated = lastUpdatedAt
		? dateFormat.format(new Date(lastUpdatedAt))
		: "لا توجد بيانات محفوظة";

	async function handleFileChange(key: UploadKey, event: ChangeEvent<HTMLInputElement>) {
		const file = event.currentTarget.files?.[0] ?? null;
		event.currentTarget.value = "";
		const nextUploads = { ...uploads, [key]: file };
		setUploads(nextUploads);
		setError("");
		setSearch("");
		setSegmentFilter("all");
		setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
		setShowOtherSegments(false);
		setPage(1);
		const currentRun = ++runId.current;

		if (!nextUploads.periodOne || !nextUploads.periodTwo || !nextUploads.segments) {
			setIsProcessing(false);
			return;
		}

		setIsProcessing(true);
		try {
			const processed = await processFiles(
				nextUploads.periodOne,
				nextUploads.periodTwo,
				nextUploads.segments,
			);
			if (currentRun === runId.current) {
				const updatedAt = new Date().toISOString();
				setResult(processed);
				setLastUpdatedAt(updatedAt);
				setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
				setShowOtherSegments(false);
				setComparison(null);
				try {
					await saveDashboard({ result: processed, updatedAt });
					const archive = await saveArchive(processed);
					setCurrentArchiveId(archive.id);
					setLastUpdatedAt(archive.createdAt);
					setArchives((current) => [archive, ...current]);
					setSelectedArchiveIds((current) => [current[0] || current[1], archive.id]);
					setArchiveStatus("connected");
				} catch (storageError) {
					setArchiveStatus("local");
					setError(`تمت المعالجة محلياً لكن تعذر حفظها في Cloudflare: ${storageError instanceof Error ? storageError.message : "خطأ غير معروف."}`);
				}
			}
		} catch (processingError) {
			if (currentRun === runId.current) {
				setError(
					processingError instanceof Error
						? processingError.message
						: "حدث خطأ غير متوقع أثناء معالجة الملفات.",
				);
			}
		} finally {
			if (currentRun === runId.current) setIsProcessing(false);
		}
	}

	async function publishLocalArchive(): Promise<void> {
		if (!result) return;
		try {
			const archive = await saveArchive(result);
			setArchives((current) => [archive, ...current]);
			setCurrentArchiveId(archive.id);
			setLastUpdatedAt(archive.createdAt);
			setArchiveStatus("connected");
			setError("");
		} catch (archiveError) {
			setError(`تعذر نشر النسخة المحلية إلى الأرشيف: ${archiveError instanceof Error ? archiveError.message : "خطأ غير معروف."}`);
		}
	}

	async function removeArchive(archive: ArchiveSummary): Promise<void> {
		if (!window.confirm(`هل تريد حذف نسخة ${dateFormat.format(new Date(archive.createdAt))} نهائياً من الأرشيف؟`)) return;
		try {
			await deleteArchive(archive.id);
			const remaining = archives.filter((item) => item.id !== archive.id);
			setArchives(remaining);
			setComparison(null);
			if (currentArchiveId === archive.id) {
				setCurrentArchiveId(remaining[0]?.id ?? null);
				if (remaining[0]) {
					const latest = await loadArchive(remaining[0].id);
					setResult(latest);
					setLastUpdatedAt(latest.createdAt);
				} else {
					setResult(null);
					setLastUpdatedAt(null);
				}
			}
			setSelectedArchiveIds([remaining[1]?.id ?? "", remaining[0]?.id ?? ""]);
			setError("");
		} catch (archiveError) {
			setError(`تعذر حذف النسخة: ${archiveError instanceof Error ? archiveError.message : "خطأ غير معروف."}`);
		}
	}

	async function compareArchives(): Promise<void> {
		const [firstId, secondId] = selectedArchiveIds;
		if (!firstId || !secondId || firstId === secondId) {
			setError("اختر نسختين مختلفتين لمقارنة أداء المندوبين.");
			return;
		}
		setIsComparing(true);
		try {
			const [first, second] = await Promise.all([loadArchive(firstId), loadArchive(secondId)]);
			const aggregate = (dataset: ArchivedDataset) => {
				const totals = new Map<string, Omit<SegmentComparison, "segment">>();
				for (const customer of dataset.rows) {
					const current = totals.get(customer.segment) ?? {
						firstBalance: 0, firstOver30: 0, firstOver45: 0, firstCommission: 0,
						secondBalance: 0, secondOver30: 0, secondOver45: 0, secondCommission: 0,
					};
					current.firstBalance += customer.total;
					current.firstOver30 += sumBucketsFrom(customer, OVER_30_BUCKET_START);
					current.firstOver45 += sumBucketsFrom(customer, OVER_45_BUCKET_START);
					current.firstCommission += sumBucketsFrom(customer, OVER_45_BUCKET_START) * COMMISSION_RATE;
					totals.set(customer.segment, current);
				}
				return totals;
			};
			const firstTotals = aggregate(first);
			const secondTotals = aggregate(second);
			const rows = [...new Set([...firstTotals.keys(), ...secondTotals.keys()])]
				.map((segment): SegmentComparison => {
					const previous = firstTotals.get(segment);
					const current = secondTotals.get(segment);
					return {
						segment,
						firstBalance: previous?.firstBalance ?? 0,
						firstOver30: previous?.firstOver30 ?? 0,
						firstOver45: previous?.firstOver45 ?? 0,
						firstCommission: previous?.firstCommission ?? 0,
						secondBalance: current?.firstBalance ?? 0,
						secondOver30: current?.firstOver30 ?? 0,
						secondOver45: current?.firstOver45 ?? 0,
						secondCommission: current?.firstCommission ?? 0,
					};
				})
				.sort((left, right) => (right.secondOver45 - right.firstOver45) - (left.secondOver45 - left.firstOver45));
			setComparison({ first, second, rows });
			setError("");
		} catch (archiveError) {
			setError(`تعذرت مقارنة النسختين: ${archiveError instanceof Error ? archiveError.message : "خطأ غير معروف."}`);
		} finally {
			setIsComparing(false);
		}
	}

	function toggleSort(column: SortColumn, bucketIndex?: number): void {
		setSort((current) => {
			const sameColumn = current.column === column && current.bucketIndex === bucketIndex;
			return {
				column,
				bucketIndex,
				direction: sameColumn && current.direction === "desc" ? "asc" : "desc",
			};
		});
		setPage(1);
	}

	function resizeColumn(columnIndex: number, event: ReactPointerEvent<HTMLButtonElement>): void {
		if (event.type === "pointerdown") {
			event.preventDefault();
			event.currentTarget.setPointerCapture(event.pointerId);
			resizeStart.current = {
				columnIndex,
				pointerX: event.clientX,
				width: columnWidths[columnIndex] ?? INITIAL_COLUMN_WIDTHS[columnIndex] ?? 125,
			};
			return;
		}
		const start = resizeStart.current;
		if (event.type === "pointerup" || event.type === "pointercancel") {
			resizeStart.current = null;
			return;
		}
		if (event.type === "pointermove" && start?.columnIndex === columnIndex) {
			const nextWidth = Math.max(
				MIN_COLUMN_WIDTH,
				start.width + start.pointerX - event.clientX,
			);
			setColumnWidths((current) =>
				current.map((width, index) => index === columnIndex ? nextWidth : width),
			);
		}
	}

	function handleResizeKey(columnIndex: number, event: KeyboardEvent<HTMLButtonElement>): void {
		if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return;
		event.preventDefault();
		const delta = event.key === "ArrowLeft" ? 10 : -10;
		setColumnWidths((current) =>
			current.map((width, index) =>
				index === columnIndex ? Math.max(MIN_COLUMN_WIDTH, width + delta) : width,
			),
		);
	}

	function renderResizeHandle(columnIndex: number, label: string) {
		return (
			<button
				type="button"
				className="column-resizer"
				aria-label={`تغيير عرض عمود ${label}`}
				onPointerDown={(event) => resizeColumn(columnIndex, event)}
				onPointerMove={(event) => resizeColumn(columnIndex, event)}
				onPointerUp={(event) => resizeColumn(columnIndex, event)}
				onPointerCancel={(event) => resizeColumn(columnIndex, event)}
				onKeyDown={(event) => handleResizeKey(columnIndex, event)}
			/>
		);
	}

	return (
		<main className="page-shell" dir="rtl">
			<header className="topbar">
				<a className="brand" href="#top" aria-label="العودة إلى البداية">
					<span className="brand-mark" aria-hidden="true">
						<svg viewBox="0 0 32 32" fill="none">
							<path d="M5 23.5 13 15l5 4.5L27 9" />
							<path d="M20 9h7v7" />
						</svg>
					</span>
					<span className="brand-name">مِيزان<span>.</span></span>
				</a>
				<div className="header-uploads" aria-label="رفع أو تحديث الملفات">
					<HeaderUpload
						title="الفترات الأولى"
						accept=".csv,text/csv"
						file={uploads.periodOne}
						disabled={isRestoring}
						onChange={(event) => void handleFileChange("periodOne", event)}
					/>
					<HeaderUpload
						title="استكمال الفترات"
						accept=".csv,text/csv"
						file={uploads.periodTwo}
						disabled={isRestoring}
						onChange={(event) => void handleFileChange("periodTwo", event)}
					/>
					<HeaderUpload
						title="ملف السيجمينت"
						accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
						file={uploads.segments}
						disabled={isRestoring}
						onChange={(event) => void handleFileChange("segments", event)}
					/>
				</div>
				<div className="header-data-info">
					<span>آخر تحديث للبيانات</span>
					<strong>{isRestoring ? "جاري الاستعادة…" : formattedLastUpdated}</strong>
				</div>
			</header>
			<div className="header-notice" role="status">
				<span className="privacy-dot" />
				{archiveStatus === "connected"
					? "أرشيف Cloudflare مشترك للعامة: أي شخص معه الرابط يمكنه القراءة والرفع والحذف."
					: archiveStatus === "checking"
						? "جاري الاتصال بأرشيف Cloudflare…"
						: "الأرشيف السحابي غير متاح؛ أي بيانات جديدة محفوظة محلياً فقط حتى عودة الاتصال."}
			</div>
			{isProcessing && (
				<div className="status-message processing" role="status">
					<i className="spinner" />
					جاري التحقق من الفترات ومطابقة الحسابات…
				</div>
			)}
			{error && (
				<div className="status-message error" role="alert">
					<span>!</span>
					{error}
				</div>
			)}

			<section className="hero" id="top">
				<div className="hero-copy">
					<div className="eyebrow"><span /> تحليل أعمار الديون</div>
					<h1>كل فتراتك،<br /><span>في ملف واحد.</span></h1>
					<p>
						ارفع تصديري Dynamics وملف السيجمينت. سنجمع الفترات غير المكررة لكل
						عميل ونتأكد أن مجموعها يطابق رصيده الفعلي.
					</p>
					<div className="hero-footnote">
						<span className="sparkle">✳</span>
						ملف Excel موحّد، جاهز للعرض
					</div>
				</div>
				<div className="hero-art" aria-hidden="true">
					<div className="art-orbit orbit-one" />
					<div className="art-orbit orbit-two" />
					<div className="art-card">
						<div className="art-card-header"><i /><i /><i /></div>
						<div className="art-card-total">9</div>
						<div className="art-card-caption">فترات عمرية في ملف واحد</div>
						<div className="art-chart">
							<span style={{ height: "36%" }} />
							<span style={{ height: "52%" }} />
							<span style={{ height: "42%" }} />
							<span style={{ height: "68%" }} />
							<span style={{ height: "56%" }} />
							<span style={{ height: "82%" }} />
							<span style={{ height: "72%" }} />
							<span style={{ height: "100%" }} />
						</div>
						<div className="art-card-footer"><span /> فترات العمر والرصيد</div>
					</div>
					<div className="art-badge">✓ مطابق</div>
					<div className="art-spark">✳</div>
				</div>
			</section>

			<section className="workflow" aria-label="خطوات المعالجة">
				<div className="workflow-step active">
					<span>١</span><strong>رفع الملفات</strong>
				</div>
				<div className={`workflow-line${result ? " complete" : ""}`} />
				<div className={`workflow-step${isProcessing || result ? " active" : ""}`}>
					<span>{isProcessing ? <i className="spinner" /> : "٢"}</span>
					<strong>دمج الفترات والسيجمينت</strong>
				</div>
				<div className={`workflow-line${result ? " complete" : ""}`} />
				<div className={`workflow-step${result ? " active" : ""}`}>
					<span>٣</span><strong>لوحة العرض وExcel</strong>
				</div>
			</section>

			{result && (
				<section className="dashboard" aria-labelledby="dashboard-title">
					<div className="dashboard-heading">
						<div>
							<p className="section-kicker">ملخص تنفيذي</p>
							<h2 id="dashboard-title">لوحة توزيع الرصيد</h2>
							<p className="dashboard-subtitle">
								دمجنا الفترات غير المكررة وربطنا السيجمينت برقم الحساب.
							</p>
						</div>
						<div className="dashboard-actions">
							<button className="button button-quiet" type="button" onClick={() => window.print()}>
								<span aria-hidden="true">⎙</span> طباعة لوحة العرض
							</button>
							<button
								className="button button-dark"
								type="button"
								onClick={() =>
									downloadExcel({
										...result,
										rows: selectedRows,
										total: selectedTotal,
										segmentMatches: selectedRows.filter((customer) => customer.segment !== "غير محدد").length,
									})
								}
							>
								<span aria-hidden="true">↓</span> تنزيل Excel المحدد
							</button>
						</div>
					</div>

					<section className="panel archive-panel" aria-labelledby="archive-title">
						<div className="panel-heading">
							<div>
								<p className="panel-kicker">Cloudflare D1 · أرشيف مشترك</p>
								<h3 id="archive-title">نسخ البيانات ومقارنة أداء المندوبين</h3>
							</div>
							{archiveStatus === "local" && result && (
								<button className="button button-dark" type="button" onClick={() => void publishLocalArchive()}>
									نشر النسخة المحلية
								</button>
							)}
						</div>
						{archives.length === 0 ? (
							<p className="archive-empty">
								لا توجد نسخ في الأرشيف بعد. ارفع الملفات الثلاثة لإنشاء أول نسخة محفوظة على Cloudflare.
							</p>
						) : (
							<>
								<div className="archive-history">
									{archives.map((archive) => (
										<div className="archive-item" key={archive.id}>
											<div>
												<strong>{dateFormat.format(new Date(archive.createdAt))}</strong>
												<span>{archive.rowCount} حساب · رصيد {formatAmount(archive.total)}</span>
											</div>
											<button type="button" onClick={() => void removeArchive(archive)}>حذف النسخة</button>
										</div>
									))}
								</div>
								<div className="comparison-controls">
									<label>
										<span>النسخة الأقدم</span>
										<select
											value={selectedArchiveIds[0]}
											onChange={(event) => setSelectedArchiveIds((current) => [event.target.value, current[1]])}
										>
											<option value="">اختر نسخة</option>
											{archives.map((archive) => <option key={archive.id} value={archive.id}>{dateFormat.format(new Date(archive.createdAt))}</option>)}
										</select>
									</label>
									<label>
										<span>النسخة الأحدث</span>
										<select
											value={selectedArchiveIds[1]}
											onChange={(event) => setSelectedArchiveIds(([first]) => [first, event.target.value])}
										>
											<option value="">اختر نسخة</option>
											{archives.map((archive) => <option key={archive.id} value={archive.id}>{dateFormat.format(new Date(archive.createdAt))}</option>)}
										</select>
									</label>
									<button
										className="button button-dark"
										type="button"
										disabled={isComparing || archives.length < 2}
										onClick={() => void compareArchives()}
									>
										{isComparing ? "جاري المقارنة…" : "مقارنة المندوبين"}
									</button>
								</div>
							</>
						)}
						{comparison && (
							<div className="comparison-results">
								<p>
									المقارنة من {dateFormat.format(new Date(comparison.first.createdAt))}
									{" إلى "}
									{dateFormat.format(new Date(comparison.second.createdAt))}
								</p>
								<div className="comparison-table-scroll">
									<table className="comparison-table">
										<thead>
											<tr>
												<th scope="col">المندوب / السيجمينت</th>
												<th scope="col">+30 سابق</th><th scope="col">+30 حالي</th>
												<th scope="col">+45 سابق</th><th scope="col">+45 حالي</th><th scope="col">فرق +45</th>
												<th scope="col">العمولة السابقة</th><th scope="col">العمولة الحالية</th><th scope="col">فرق العمولة</th>
											</tr>
										</thead>
										<tbody>
											{comparison.rows.map((row) => (
												<tr key={row.segment}>
													<td>{row.segment}</td>
													<td className="numeric-cell">{formatAmount(row.firstOver30)}</td>
													<td className="numeric-cell">{formatAmount(row.secondOver30)}</td>
													<td className="numeric-cell">{formatAmount(row.firstOver45)}</td>
													<td className="numeric-cell">{formatAmount(row.secondOver45)}</td>
													<td className="numeric-cell">{formatAmount(row.secondOver45 - row.firstOver45)}</td>
													<td className="numeric-cell">{formatAmount(row.firstCommission)}</td>
													<td className="numeric-cell">{formatAmount(row.secondCommission)}</td>
													<td className="numeric-cell commission-value">{formatAmount(row.secondCommission - row.firstCommission)}</td>
												</tr>
											))}
										</tbody>
									</table>
								</div>
							</div>
						)}
					</section>

					<section className="owner-filter-panel" aria-labelledby="owner-filter-title">
						<div className="owner-filter-heading">
							<div>
								<p className="panel-kicker">إظهار الحسابات التابعة لـ</p>
								<h3 id="owner-filter-title">اختيار الأشخاص المهمين</h3>
							</div>
							<div className="owner-filter-actions">
								<button
									type="button"
									onClick={() => {
										setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
										setShowOtherSegments(false);
										setPage(1);
									}}
								>
									تحديد المهمين
								</button>
								<button
									type="button"
									onClick={() => {
										setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
										setShowOtherSegments(true);
										setPage(1);
									}}
								>
									إظهار الكل
								</button>
							</div>
						</div>
						<div className="owner-checkbox-grid">
							{OWNER_FILTERS.map((owner) => {
								const checked = selectedOwners.has(owner.id);
								const count = ownerCounts.get(owner.id) ?? 0;
								return (
									<label className={`owner-checkbox${checked ? " checked" : ""}`} key={owner.id}>
										<input
											type="checkbox"
											checked={checked}
											onChange={(event) => {
												setSelectedOwners((current) => {
													const next = new Set(current);
													if (event.target.checked) next.add(owner.id);
													else next.delete(owner.id);
													return next;
												});
												setSegmentFilter("all");
												setPage(1);
											}}
										/>
										<span className="custom-checkbox" aria-hidden="true">{checked ? "✓" : ""}</span>
										<span className="owner-checkbox-label">
											<strong>{owner.name}</strong>
											<small>{count > 0 ? `${count} حساب` : "غير موجود في ملف السيجمينت"}</small>
										</span>
									</label>
								);
							})}
							<label className={`owner-checkbox other-owner-checkbox${showOtherSegments ? " checked" : ""}`}>
								<input
									type="checkbox"
									checked={showOtherSegments}
									onChange={(event) => {
										setShowOtherSegments(event.target.checked);
										setSegmentFilter("all");
										setPage(1);
									}}
								/>
								<span className="custom-checkbox" aria-hidden="true">{showOtherSegments ? "✓" : ""}</span>
								<span className="owner-checkbox-label">
									<strong>بقية السيجمينتات</strong>
									<small>{otherSegmentCount} حساب</small>
								</span>
							</label>
						</div>
						<div className="owner-filter-footer" role="status">
							<span>ظاهر: {selectedRows.length} من {result.rows.length} حساب</span>
						</div>
					</section>

					<div className="metric-grid">
						<article className="metric-card metric-primary">
							<div className="metric-label"><span className="legend-dot dot-second" /> إجمالي الرصيد المعتمد</div>
							<strong>{formatAmount(selectedTotal)}</strong>
							<span className="metric-note">مجموع أرصدة الأشخاص والسيجمينتات المحددة</span>
						</article>
						<article className="metric-card">
							<div className="metric-label">الحسابات</div>
							<strong>{selectedRows.length}</strong>
							<span className="metric-note">من أصل {result.rows.length} حساب</span>
						</article>
						<article className="metric-card">
							<div className="metric-label">الفترات الموحّدة</div>
							<strong>{result.bucketNames.length}</strong>
							<span className="metric-note">بعد حذف الفترات المشتركة والتجميعية</span>
						</article>
						<article className="metric-card">
							<div className="metric-label">السيجمينت</div>
							<strong>{selectedRows.filter((customer) => customer.segment !== "غير محدد").length} / {selectedRows.length}</strong>
							<span className="metric-note">{segments.length} تصنيفاً · المفقود يظهر «غير محدد»</span>
						</article>
						<article className="metric-card overdue-metric">
							<div className="metric-label">المتأخرات فوق 30 يوم</div>
							<strong>{formatAmount(selectedOver30)}</strong>
							<span className="metric-note">{formatPercent(selectedOver30, selectedTotal)} من الرصيد الفعلي</span>
						</article>
						<article className="metric-card overdue-metric">
							<div className="metric-label">المتأخرات فوق 45 يوم</div>
							<strong>{formatAmount(selectedOver45)}</strong>
							<span className="metric-note">{formatPercent(selectedOver45, selectedTotal)} من الرصيد الفعلي</span>
						</article>
						<article className="metric-card commission-metric">
							<div className="metric-label">العمولة المفقودة للمندوب · 1٪</div>
							<strong>{formatAmount(selectedOver45 * COMMISSION_RATE)}</strong>
							<span className="metric-note">محسوبة على إجمالي الرصيد المتأخر فوق 45 يوماً</span>
						</article>
					</div>

					<div className="chart-grid">
						<article className="panel age-panel">
							<div className="panel-heading">
								<div>
									<p className="panel-kicker">مجموع الأرصدة حسب الفترات</p>
									<h3>توزيع إجمالي الرصيد</h3>
								</div>
								<span className="panel-icon">▥</span>
							</div>
							<div className="age-chart">
								{bucketTotals.map((bucket, index) => (
									<div className="age-row" key={`${index}-${bucket.name}`}>
										<div className="age-label" title={bucket.name}>
											<span>{bucket.name}</span>
											<strong>{formatAmount(bucket.amount)}</strong>
										</div>
										<div className="bar-track">
											<div
												className={`bar-fill age-fill segment-color-${index % 5}`}
												style={{ width: `${Math.max(0, (bucket.amount / Math.max(maxBucket, 1)) * 100)}%` }}
											/>
										</div>
									</div>
								))}
							</div>
							<div className="chart-footnote">
								<span>مجموع الفترات</span>
								<strong>{formatAmount(selectedTotal)}</strong>
							</div>
						</article>

						<article className="panel segment-panel">
							<div className="panel-heading">
								<div>
									<p className="panel-kicker">حسب ملف السيجمينت</p>
									<h3>الرصيد حسب التصنيف</h3>
								</div>
								<span className="panel-period">{segments.length} سيجمينت</span>
							</div>
							<div className="segment-chart">
								{segments.slice(0, 8).map(([segment, amount], index) => (
									<div className="segment-row" key={segment}>
										<div className="segment-label" title={segment}>
											<span className={`segment-dot segment-color-${index % 5}`} />
											<span>{segment}</span>
											<strong>{formatAmount(amount)}</strong>
										</div>
										<div className="bar-track segment-track">
											<div
												className={`bar-fill segment-fill segment-color-${index % 5}`}
												style={{ width: `${Math.max(0, (amount / Math.max(maxSegment, 1)) * 100)}%` }}
											/>
										</div>
									</div>
								))}
								{otherSegments.length > 0 && (
									<div className="segment-row">
										<div className="segment-label">
											<span className="segment-dot segment-color-other" />
											<span>بقية السيجمينتات ({otherSegments.length})</span>
											<strong>{formatAmount(otherSegmentTotal)}</strong>
										</div>
										<div className="bar-track segment-track">
											<div
												className="bar-fill segment-fill segment-color-other"
												style={{ width: `${Math.max(0, (otherSegmentTotal / Math.max(maxSegment, 1)) * 100)}%` }}
											/>
										</div>
									</div>
								)}
							</div>
							<div className="chart-footnote">
								<span>الأرصدة غير المصنفة</span>
								<strong>{selectedRows.filter((customer) => customer.segment === "غير محدد").length} حساب</strong>
							</div>
						</article>
					</div>

					<section className="panel commission-panel" aria-labelledby="commission-title">
						<div className="panel-heading">
							<div>
								<p className="panel-kicker">الرصيد المتأخر × نسبة العمولة</p>
								<h3 id="commission-title">العمولة المفقودة حسب المندوب</h3>
							</div>
							<span className="panel-period">1٪ من رصيد +45 يوم</span>
						</div>
						<div className="commission-table-scroll">
							<table className="commission-table">
								<thead>
									<tr>
										<th scope="col">المندوب / السيجمينت</th>
										<th scope="col">الحسابات</th>
										<th scope="col">الرصيد الفعلي</th>
										<th scope="col">فوق 30 يوم</th>
										<th scope="col">النسبة</th>
										<th scope="col">فوق 45 يوم</th>
										<th scope="col">النسبة</th>
										<th scope="col">العمولة المفقودة</th>
									</tr>
								</thead>
								<tbody>
									{commissionBySegment.map(([segment, totals]) => (
										<tr key={segment}>
											<td><span className="segment-tag">{segment}</span></td>
											<td className="numeric-cell">{totals.count}</td>
											<td className="numeric-cell">{formatAmount(totals.total)}</td>
											<td className="numeric-cell">{formatAmount(totals.over30)}</td>
											<td className="numeric-cell">{formatPercent(totals.over30, totals.total)}</td>
											<td className="numeric-cell">{formatAmount(totals.over45)}</td>
											<td className="numeric-cell">{formatPercent(totals.over45, totals.total)}</td>
											<td className="numeric-cell commission-value">{formatAmount(totals.over45 * COMMISSION_RATE)}</td>
										</tr>
									))}
									{commissionBySegment.length === 0 && (
										<tr><td colSpan={8} className="commission-empty">لا توجد حسابات ضمن الأشخاص والسيجمينتات المحددة.</td></tr>
									)}
								</tbody>
								{commissionBySegment.length > 0 && (
									<tfoot>
										<tr>
											<th scope="row">الإجمالي</th>
											<td className="numeric-cell">{selectedRows.length}</td>
											<td className="numeric-cell">{formatAmount(selectedTotal)}</td>
											<td className="numeric-cell">{formatAmount(selectedOver30)}</td>
											<td className="numeric-cell">{formatPercent(selectedOver30, selectedTotal)}</td>
											<td className="numeric-cell">{formatAmount(selectedOver45)}</td>
											<td className="numeric-cell">{formatPercent(selectedOver45, selectedTotal)}</td>
											<td className="numeric-cell commission-value">{formatAmount(selectedOver45 * COMMISSION_RATE)}</td>
										</tr>
									</tfoot>
								)}
							</table>
						</div>
					</section>

					<section className="panel accounts-panel" aria-labelledby="accounts-title">
						<div className="accounts-heading">
							<div>
								<p className="panel-kicker">كل الحسابات والفترات</p>
								<h3 id="accounts-title">تفاصيل الرصيد الموحّد</h3>
							</div>
							<div className="table-controls">
								<label className="search-box">
									<span aria-hidden="true">⌕</span>
									<input
										type="search"
										value={search}
										onChange={(event) => {
											setSearch(event.target.value);
											setPage(1);
										}}
										placeholder="ابحث عن حساب أو اسم..."
										aria-label="ابحث عن حساب أو اسم"
									/>
								</label>
								<select
									value={segmentFilter}
									onChange={(event) => {
										setSegmentFilter(event.target.value);
										setPage(1);
									}}
									aria-label="تصفية حسب السيجمينت"
								>
									<option value="all">كل السيجمينتات</option>
									{segments.map(([segment]) => <option key={segment} value={segment}>{segment}</option>)}
								</select>
							</div>
						</div>
						<div className="table-summary">
							عرض <strong>{visibleRows.length === 0 ? 0 : (page - 1) * PAGE_SIZE + 1}–{Math.min(page * PAGE_SIZE, visibleRows.length)}</strong> من <strong>{visibleRows.length}</strong> نتيجة
							<span> · </span>
							{selectedRows.length} حساب محدد · إجمالي {formatAmount(selectedTotal)}
						</div>
						<div className="table-scroll">
							<table
								className="accounts-table"
								style={{ width: `${columnWidths.reduce((sum, width) => sum + width, 0)}px` }}
							>
								<colgroup>
									{columnWidths.map((width, index) => <col key={index} style={{ width: `${width}px` }} />)}
								</colgroup>
								<thead>
									<tr>
										<th scope="col">الحساب / العميل{renderResizeHandle(0, "الحساب")}</th>
										<th scope="col" aria-sort={sort.column === "total" ? (sort.direction === "desc" ? "descending" : "ascending") : "none"}>
											<button className="sort-button" type="button" onClick={() => toggleSort("total")}>
												الرصيد الفعلي{sort.column === "total" ? (sort.direction === "desc" ? " ↓" : " ↑") : ""}
											</button>
											{renderResizeHandle(1, "الرصيد الفعلي")}
										</th>
										<th scope="col">السيجمينت{renderResizeHandle(2, "السيجمينت")}</th>
										{result.bucketNames.map((name, index) => (
											<th
												scope="col"
												key={`${index}-${name}`}
												aria-sort={sort.column === "bucket" && sort.bucketIndex === index ? (sort.direction === "desc" ? "descending" : "ascending") : "none"}
											>
												<button className="sort-button" type="button" onClick={() => toggleSort("bucket", index)}>
													{name}{sort.column === "bucket" && sort.bucketIndex === index ? (sort.direction === "desc" ? " ↓" : " ↑") : ""}
												</button>
												{renderResizeHandle(index + 3, name)}
											</th>
										))}
										<th scope="col">عمولة 1٪{renderResizeHandle(result.bucketNames.length + 3, "العمولة")}</th>
									</tr>
								</thead>
								<tbody>
									{pageRows.map((customer) => (
										<tr key={customer.account}>
											<td>
												<strong className="account-name">{customer.name || "—"}</strong>
												<span className="account-id">{customer.account}</span>
											</td>
											<td className="numeric-cell total-cell">{formatAmount(customer.total)}</td>
											<td><span className="segment-tag">{customer.segment}</span></td>
											{customer.buckets.map((amount, index) => (
												<td className="numeric-cell" key={`${customer.account}-${index}`}>{formatAmount(amount)}</td>
											))}
											<td className="numeric-cell commission-value">
												{formatAmount(sumBucketsFrom(customer, OVER_45_BUCKET_START) * COMMISSION_RATE)}
											</td>
										</tr>
									))}
								</tbody>
							</table>
							{visibleRows.length === 0 && <div className="empty-table">لا توجد نتائج تطابق البحث.</div>}
						</div>
						{visibleRows.length > PAGE_SIZE && (
							<div className="table-pagination">
								<button
									type="button"
									disabled={page <= 1}
									onClick={() => setPage((currentPage) => Math.max(1, currentPage - 1))}
								>
									السابق
								</button>
								<span>صفحة {page} من {pageCount}</span>
								<button
									type="button"
									disabled={page >= pageCount}
									onClick={() => setPage((currentPage) => Math.min(pageCount, currentPage + 1))}
								>
									التالي
								</button>
							</div>
						)}
					</section>

					<div className="result-note">
						<span aria-hidden="true">i</span>
						نعرض فقط الأشخاص والسيجمينتات المحددة؛ أزل علامة الصح لإخفاء حساباتهم من الجدول والملخص والرسوم وملف Excel.
						{result.segmentMatches < result.rows.length && (
							<span> لم نجد سيجمينتاً لـ {result.rows.length - result.segmentMatches} حساب؛ تظهر كـ «غير محدد» ضمن بقية السيجمينتات.</span>
						)}
					</div>
				</section>
			)}

			<footer className="page-footer">
				<span>مِيزان<span className="footer-dot">.</span> تحليل الذمم ببساطة</span>
				<span>معالجة محلية · بدون رفع بياناتك إلى خادم</span>
			</footer>
		</main>
	);
}

export default App;
