import { useMemo, useRef, useState, type ChangeEvent, type KeyboardEvent, type PointerEvent as ReactPointerEvent } from "react";
import { createExcelFile } from "./excel-export";
import { processFiles, type ProcessingResult } from "./processor";
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
const INITIAL_COLUMN_WIDTHS = [260, 145, 150, ...Array(9).fill(125)];
const amountFormat = new Intl.NumberFormat("en-US", {
	minimumFractionDigits: 2,
	maximumFractionDigits: 2,
});
type SortColumn = "total" | "bucket";
type SortDirection = "asc" | "desc";
interface TableSort {
	column: SortColumn;
	bucketIndex?: number;
	direction: SortDirection;
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

function UploadCard({
	number,
	title,
	description,
	accept,
	file,
	onChange,
}: {
	number: string;
	title: string;
	description: string;
	accept: string;
	file: File | null;
	onChange: (event: ChangeEvent<HTMLInputElement>) => void;
}) {
	return (
		<div className={`upload-card${file ? " has-file" : ""}`}>
			<div className="upload-card-top">
				<span className="upload-number">{number}</span>
				<span className="upload-state" aria-label={file ? "تم اختيار الملف" : "بانتظار الملف"}>
					{file ? "✓" : "↑"}
				</span>
			</div>
			<div className="upload-copy">
				<h3>{title}</h3>
				<p>{description}</p>
			</div>
			<label className="file-picker">
				<input type="file" accept={accept} onChange={onChange} />
				<span>{file ? "تغيير الملف" : "اختيار ملف"}</span>
			</label>
			<div className={`file-name${file ? " selected" : ""}`} title={file?.name}>
				{file ? file.name : "لم يتم اختيار ملف"}
			</div>
		</div>
	);
}

function App() {
	const [uploads, setUploads] = useState<Uploads>(EMPTY_UPLOADS);
	const [result, setResult] = useState<ProcessingResult | null>(null);
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

	const pageCount = Math.max(1, Math.ceil(visibleRows.length / PAGE_SIZE));
	const pageRows = visibleRows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);
	const maxBucket = Math.max(0, ...bucketTotals.map((bucket) => bucket.amount));
	const maxSegment = Math.max(0, ...segments.slice(0, 8).map(([, amount]) => amount));
	const otherSegments = segments.slice(8);
	const otherSegmentTotal = otherSegments.reduce((sum, [, amount]) => sum + amount, 0);

	async function handleFileChange(key: UploadKey, event: ChangeEvent<HTMLInputElement>) {
		const file = event.currentTarget.files?.[0] ?? null;
		event.currentTarget.value = "";
		const nextUploads = { ...uploads, [key]: file };
		setUploads(nextUploads);
		setResult(null);
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
				setResult(processed);
				setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
				setShowOtherSegments(false);
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

	function reset(): void {
		runId.current += 1;
		setUploads(EMPTY_UPLOADS);
		setResult(null);
		setError("");
		setIsProcessing(false);
		setSearch("");
		setSegmentFilter("all");
		setSelectedOwners(new Set());
		setShowOtherSegments(false);
		setPage(1);
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
				<div className="privacy-note">
					<span className="privacy-dot" />
					تتم المعالجة على جهازك فقط
				</div>
			</header>

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

			<section className="upload-section" aria-labelledby="upload-title">
				<div className="section-heading">
					<div>
						<p className="section-kicker">ابدأ من هنا</p>
						<h2 id="upload-title">ملفات التحليل</h2>
					</div>
					<span className="file-count">ملفان CSV · ملف سيجمينت Excel</span>
				</div>
				<div className="upload-grid">
					<UploadCard
						number="01"
						title="الفترات الأولى"
						description="الأعمار الحالية وحتى أقل من 37 يوماً"
						accept=".csv,text/csv"
						file={uploads.periodOne}
						onChange={(event) => void handleFileChange("periodOne", event)}
					/>
					<UploadCard
						number="02"
						title="استكمال الفترات"
						description="الفترات الإضافية بعد الأعمار المشتركة"
						accept=".csv,text/csv"
						file={uploads.periodTwo}
						onChange={(event) => void handleFileChange("periodTwo", event)}
					/>
					<UploadCard
						number="03"
						title="ملف السيجمينت"
						description="ملف Excel يحتوي Customer account وSegment"
						accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
						file={uploads.segments}
						onChange={(event) => void handleFileChange("segments", event)}
					/>
				</div>
				<div className="upload-footer">
					<div className="upload-hint">
						<span className="lock-icon" aria-hidden="true">⌑</span>
						الملفات لا تُرفع إلى خادم ولا تُحفظ بعد إغلاق الصفحة.
					</div>
					{(uploads.periodOne || uploads.periodTwo || uploads.segments) && (
						<button className="text-button" type="button" onClick={reset}>
							مسح الملفات
						</button>
					)}
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
