import { useEffect, useMemo, useRef, useState, type ChangeEvent, type DragEvent } from "react";
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

const OVER_30_BUCKET_START = 5;
const OVER_45_BUCKET_START = 7;
const AGING_CHART_COLORS = ["#3d9272", "#59a78c", "#78b89f", "#96c8b2", "#b1d4c3", "#e6c84f", "#efb83f", "#cc7654", "#a94f4f"];
const COMMISSION_RATE = 0.01;
const amountFormat = new Intl.NumberFormat("en-US", {
	minimumFractionDigits: 2,
	maximumFractionDigits: 2,
});
const trendAxisAmountFormat = new Intl.NumberFormat("en-US", {
	notation: "compact",
	maximumFractionDigits: 1,
});
const percentFormat = new Intl.NumberFormat("ar", {
	minimumFractionDigits: 0,
	maximumFractionDigits: 2,
});
const reportDateFormat = new Intl.DateTimeFormat("ar", { dateStyle: "medium", timeZone: "UTC" });
const trendDateFormat = new Intl.DateTimeFormat("en-GB", { day: "2-digit", month: "2-digit", timeZone: "UTC" });
type SortDirection = "asc" | "desc";
interface HeatmapSort {
	column: "customer" | "segment" | "total" | "bucket";
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

interface ArchiveSnapshotRows {
	id: string;
	createdAt: string;
	reportDate: string | null;
	rows: ProcessingResult["rows"];
}

function archivedResult(dataset: ArchivedDataset): ProcessingResult {
	return {
		bucketNames: dataset.bucketNames,
		rows: dataset.rows,
		total: dataset.total,
		segmentMatches: dataset.segmentMatches,
		reportDate: dataset.reportDate ?? "",
	};
}

interface RepresentativeMetrics {
	id: string;
	name: string;
	rows: ProcessingResult["rows"];
	total: number;
	under21: number;
	under45: number;
	over45: number;
	over45Accounts: number;
	previousTotal: number | null;
	changeAmount: number | null;
	changePercent: number | null;
}

type DateRangePreset = "today" | "yesterday" | "last-7-days" | "last-30-days" | "last-12-months" | "custom";
type DateRange = { start: string; end: string };
type RepresentativeComparisonMetric = "balance" | "over30" | "over45" | "commission";
type DashboardWidgetId =
	| "kpi-total"
	| "kpi-under21"
	| "kpi-under30"
	| "kpi-30-45"
	| "kpi-over45"
	| "kpi-commission"
	| "trend"
	| "aging-pie"
	| "owner-comparison"
	| "balance-distribution"
	| "over45-concentration"
	| "owner-balances"
	| "heatmap";

const TREND_SNAPSHOT_LIMIT = 12;
const REPRESENTATIVE_COMPARISON_COLORS = ["#176f5e", "#d47a32"];
const DASHBOARD_WIDGET_STORAGE_KEY = "mizan-dashboard-widgets-v1";
const DASHBOARD_WIDGET_ORDER_STORAGE_KEY = "mizan-dashboard-widget-order-v1";
const DEFAULT_VISIBLE_WIDGETS: Record<DashboardWidgetId, boolean> = {
	"kpi-total": true,
	"kpi-under21": true,
	"kpi-under30": true,
	"kpi-30-45": true,
	"kpi-over45": true,
	"kpi-commission": true,
	trend: true,
	"aging-pie": true,
	"owner-comparison": true,
	"balance-distribution": true,
	"over45-concentration": true,
	"owner-balances": true,
	heatmap: true,
};
const DASHBOARD_WIDGET_GROUPS: {
	title: string;
	items: { id: DashboardWidgetId; label: string }[];
}[] = [
	{
		title: "بطاقات المؤشرات",
		items: [
			{ id: "kpi-total", label: "إجمالي الرصيد" },
			{ id: "kpi-under21", label: "الرصيد حتى 21 يوماً" },
			{ id: "kpi-under30", label: "الرصيد حتى 30 يوماً" },
			{ id: "kpi-30-45", label: "الرصيد بين 30–45 يوماً" },
			{ id: "kpi-over45", label: "الرصيد فوق 45 يوماً" },
			{ id: "kpi-commission", label: "العمولة التقديرية على +45" },
		],
	},
	{
		title: "الرسوم والتحليلات",
		items: [
			{ id: "trend", label: "مخطط خطي لاتجاه الرصيد" },
			{ id: "aging-pie", label: "مخطط دائري لتوزيع الفترات" },
			{ id: "owner-comparison", label: "مخطط مقارنة مندوبين" },
			{ id: "balance-distribution", label: "توزيع الرصيد حسب الفترات" },
			{ id: "over45-concentration", label: "تركيز متأخرات +45" },
			{ id: "owner-balances", label: "الرصيد حسب المندوب" },
			{ id: "heatmap", label: "الخريطة الحرارية" },
		],
	},
];
const DASHBOARD_WIDGET_IDS = DASHBOARD_WIDGET_GROUPS.flatMap((group) => group.items.map(({ id }) => id));
const DAILY_SALES_WORKBOOK_URL = "https://onedrive.live.com/:x:/g/personal/f09e9d11c43860ca/IQCXVsvY6FTdT7aJTJrTK_5sAU5EcNhqhSl1dcfuzI3XBiI?rtime=enSWX2cn30g&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy9mMDllOWQxMWM0Mzg2MGNhL0lRQ1hWc3ZZNkZUZFQ3YUpUSnJUS181c0FVNUVjTmhxaFNsMWRjZnV6STNYQmlJP2U9QmN4dElD";
type AppTab = "dashboard" | "daily-sales";
type DailySalesSheet = "المبيعات اليومية" | "متابعة الهدف";

function dailySalesEmbedUrl(sheet: DailySalesSheet, refresh: number): string {
	const url = new URL(DAILY_SALES_WORKBOOK_URL);
	url.searchParams.set("action", "embedview");
	url.searchParams.set("wdSheet", sheet);
	url.searchParams.set("wdAllowInteractivity", "True");
	url.searchParams.set("wdHideGridlines", "True");
	url.searchParams.set("refresh", String(refresh));
	return url.toString();
}

function isDateRangePreset(value: string): value is DateRangePreset {
	return value === "today" || value === "yesterday" || value === "last-7-days" ||
		value === "last-30-days" || value === "last-12-months" || value === "custom";
}

function isRepresentativeComparisonMetric(value: string): value is RepresentativeComparisonMetric {
	return value === "balance" || value === "over30" || value === "over45" || value === "commission";
}

function isDashboardWidgetId(value: string): value is DashboardWidgetId {
	return DASHBOARD_WIDGET_IDS.some((id) => id === value);
}

function readDashboardWidgetPreferences(): Record<DashboardWidgetId, boolean> {
	const saved = window.localStorage.getItem(DASHBOARD_WIDGET_STORAGE_KEY);
	if (!saved) return { ...DEFAULT_VISIBLE_WIDGETS };
	const parsed: unknown = JSON.parse(saved);
	if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
		throw new Error("تعذر قراءة تفضيلات تخصيص لوحة البيانات.");
	}
	const preferences = { ...DEFAULT_VISIBLE_WIDGETS };
	for (const group of DASHBOARD_WIDGET_GROUPS) {
		for (const { id } of group.items) {
			const value = (parsed as Record<string, unknown>)[id];
			if (value !== undefined && typeof value !== "boolean") {
				throw new Error("تفضيلات تخصيص لوحة البيانات غير صالحة.");
			}
			if (typeof value === "boolean") preferences[id] = value;
		}
	}
	return preferences;
}

function readDashboardWidgetOrder(): DashboardWidgetId[] {
	const saved = window.localStorage.getItem(DASHBOARD_WIDGET_ORDER_STORAGE_KEY);
	if (!saved) return [...DASHBOARD_WIDGET_IDS];
	const parsed: unknown = JSON.parse(saved);
	if (!Array.isArray(parsed) || parsed.some((id) => typeof id !== "string")) {
		throw new Error("تعذر قراءة ترتيب عناصر لوحة البيانات.");
	}
	const savedIds = parsed.filter((id): id is DashboardWidgetId => typeof id === "string" && isDashboardWidgetId(id));
	return [...new Set([...savedIds, ...DASHBOARD_WIDGET_IDS])];
}

function toDateInputValue(date: Date): string {
	const year = date.getFullYear();
	const month = String(date.getMonth() + 1).padStart(2, "0");
	const day = String(date.getDate()).padStart(2, "0");
	return `${year}-${month}-${day}`;
}

function getDateRange(preset: Exclude<DateRangePreset, "custom">, today = new Date()): DateRange {
	const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());
	const start = new Date(end);
	switch (preset) {
		case "yesterday":
			start.setDate(start.getDate() - 1);
			end.setDate(end.getDate() - 1);
			break;
		case "last-7-days":
			start.setDate(start.getDate() - 6);
			break;
		case "last-30-days":
			start.setDate(start.getDate() - 29);
			break;
		case "last-12-months":
			start.setFullYear(start.getFullYear() - 1);
			break;
		case "today":
			break;
	}
	return { start: toDateInputValue(start), end: toDateInputValue(end) };
}

function valueForRepresentativeMetric(
	rows: ProcessingResult["rows"],
	ownerId: string,
	metric: RepresentativeComparisonMetric,
): number {
	const representativeRows = rows.filter((customer) => ownerForSegment(customer.segment) === ownerId);
	switch (metric) {
		case "balance":
			return totalBalance(representativeRows);
		case "over30":
			return representativeRows.reduce((sum, customer) => sum + sumBucketsFrom(customer, OVER_30_BUCKET_START), 0);
		case "over45": {
			return representativeRows.reduce((sum, customer) => sum + sumBucketsFrom(customer, OVER_45_BUCKET_START), 0);
		}
		case "commission":
			return representativeRows.reduce((sum, customer) => sum + sumBucketsFrom(customer, OVER_45_BUCKET_START) * COMMISSION_RATE, 0);
	}
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

function formatReportDate(reportDate: string | null | undefined): string {
	if (reportDate && /^\d{4}-\d{2}-\d{2}$/.test(reportDate)) {
		return reportDateFormat.format(new Date(`${reportDate}T00:00:00Z`));
	}
	return "تاريخ الرصيد غير متوفر";
}

function formatTrendDate(reportDate: string | null | undefined): string {
	if (reportDate && /^\d{4}-\d{2}-\d{2}$/.test(reportDate)) {
		return trendDateFormat.format(new Date(`${reportDate}T00:00:00Z`));
	}
	return "—";
}

function archiveOptionLabel(archive: ArchiveSummary): string {
	return formatReportDate(archive.reportDate);
}

function compareByBalanceDate(
	first: Pick<ArchiveSummary, "reportDate" | "createdAt">,
	second: Pick<ArchiveSummary, "reportDate" | "createdAt">,
): number {
	return (second.reportDate ?? second.createdAt.slice(0, 10)).localeCompare(first.reportDate ?? first.createdAt.slice(0, 10)) ||
		second.createdAt.localeCompare(first.createdAt);
}

async function loadLatestArchivePair(archiveList: ArchiveSummary[]): Promise<{
	latest: ArchivedDataset;
	previous: ArchivedDataset | null;
} | null> {
	if (archiveList.length === 0) return null;
	const [latest, previous] = await Promise.all([
		loadArchive(archiveList[0].id),
		archiveList[1] ? loadArchive(archiveList[1].id) : Promise.resolve(null),
	]);
	return { latest, previous };
}

function sumBucketsFrom(customer: ProcessingResult["rows"][number], startIndex: number): number {
	return customer.buckets.slice(startIndex).reduce((sum, amount) => sum + amount, 0);
}

function sumBucketsThrough(customer: ProcessingResult["rows"][number], endIndex: number): number {
	return customer.buckets.slice(0, endIndex + 1).reduce((sum, amount) => sum + amount, 0);
}

function amountBetweenBuckets(customer: ProcessingResult["rows"][number], startIndex: number, endIndex: number): number {
	return customer.buckets.slice(startIndex, endIndex + 1).reduce((sum, amount) => sum + amount, 0);
}

function totalBalance(rows: ProcessingResult["rows"]): number {
	return rows.reduce((sum, customer) => sum + customer.total, 0);
}

function rowsForRepresentativeView(
	rows: ProcessingResult["rows"],
	ownerId: string,
	selectedOwnerIds: Set<string>,
	includeOtherSegments: boolean,
): ProcessingResult["rows"] {
	if (ownerId !== "all") {
		return rows.filter((customer) => ownerForSegment(customer.segment) === ownerId);
	}
	return rows.filter((customer) => {
		const customerOwnerId = ownerForSegment(customer.segment);
		return customerOwnerId ? selectedOwnerIds.has(customerOwnerId) : includeOtherSegments;
	});
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
	const [activeTab, setActiveTab] = useState<AppTab>("dashboard");
	const [dailySalesSheet, setDailySalesSheet] = useState<DailySalesSheet>("المبيعات اليومية");
	const [dailySalesRefresh, setDailySalesRefresh] = useState(0);
	const [uploads, setUploads] = useState<Uploads>(EMPTY_UPLOADS);
	const [result, setResult] = useState<ProcessingResult | null>(null);
	const [lastUpdatedAt, setLastUpdatedAt] = useState<string | null>(null);
	const [currentArchiveId, setCurrentArchiveId] = useState<string | null>(null);
	const [previousSnapshot, setPreviousSnapshot] = useState<ProcessingResult["rows"] | null>(null);
	const [historySnapshots, setHistorySnapshots] = useState<ArchiveSnapshotRows[]>([]);
	const [archives, setArchives] = useState<ArchiveSummary[]>([]);
	const [archiveStatus, setArchiveStatus] = useState<"checking" | "connected" | "local">("checking");
	const [selectedArchiveIds, setSelectedArchiveIds] = useState<[string, string]>(["", ""]);
	const [dateRangePreset, setDateRangePreset] = useState<DateRangePreset>("last-12-months");
	const [appliedDateRange, setAppliedDateRange] = useState<DateRange>(() => getDateRange("last-12-months"));
	const [customDateRange, setCustomDateRange] = useState<DateRange>(() => getDateRange("last-12-months"));
	const [isDateRangeLoading, setIsDateRangeLoading] = useState(false);
	const [comparisonOwnerIds, setComparisonOwnerIds] = useState<[string, string]>([
		OWNER_FILTERS[0].id,
		OWNER_FILTERS[1].id,
	]);
	const [representativeComparisonMetric, setRepresentativeComparisonMetric] =
		useState<RepresentativeComparisonMetric>("balance");
	const [visibleWidgets, setVisibleWidgets] = useState<Record<DashboardWidgetId, boolean>>(
		() => ({ ...DEFAULT_VISIBLE_WIDGETS }),
	);
	const [widgetOrder, setWidgetOrder] = useState<DashboardWidgetId[]>(() => [...DASHBOARD_WIDGET_IDS]);
	const [draggedWidgetId, setDraggedWidgetId] = useState<DashboardWidgetId | null>(null);
	const [isWidgetSettingsOpen, setIsWidgetSettingsOpen] = useState(false);
	const [comparison, setComparison] = useState<{
		first: ArchivedDataset;
		second: ArchivedDataset;
		trendSnapshots: ArchiveSnapshotRows[];
		rows: SegmentComparison[];
	} | null>(null);
	const [isComparing, setIsComparing] = useState(false);
	const [isRestoring, setIsRestoring] = useState(true);
	const [isProcessing, setIsProcessing] = useState(false);
	const [error, setError] = useState("");
	const [activeOwnerId, setActiveOwnerId] = useState<string>("all");
	const [selectedOwners, setSelectedOwners] = useState<Set<string>>(
		() => new Set(OWNER_FILTERS.map((owner) => owner.id)),
	);
	const [showOtherSegments, setShowOtherSegments] = useState(false);
	const [heatmapSort, setHeatmapSort] = useState<HeatmapSort>({ column: "total", direction: "desc" });
	const runId = useRef(0);
	const dateRangeRequestId = useRef(0);

	useEffect(() => {
		try {
			setVisibleWidgets(readDashboardWidgetPreferences());
			setWidgetOrder(readDashboardWidgetOrder());
		} catch (preferenceError) {
			setError(`تعذر تحميل تفضيلات تخصيص اللوحة: ${preferenceError instanceof Error ? preferenceError.message : "خطأ غير معروف."}`);
		}
	}, []);

	function updateWidgetVisibility(widgetId: DashboardWidgetId, visible: boolean): void {
		const updated = { ...visibleWidgets, [widgetId]: visible };
		setVisibleWidgets(updated);
		try {
			window.localStorage.setItem(DASHBOARD_WIDGET_STORAGE_KEY, JSON.stringify(updated));
			setError("");
		} catch (preferenceError) {
			setError(`تعذر حفظ تفضيلات تخصيص اللوحة على هذا الجهاز: ${preferenceError instanceof Error ? preferenceError.message : "خطأ غير معروف."}`);
		}
	}

	function resetWidgetVisibility(): void {
		const defaults = { ...DEFAULT_VISIBLE_WIDGETS };
		setVisibleWidgets(defaults);
		try {
			window.localStorage.setItem(DASHBOARD_WIDGET_STORAGE_KEY, JSON.stringify(defaults));
			setError("");
		} catch (preferenceError) {
			setError(`تعذر حفظ تفضيلات تخصيص اللوحة على هذا الجهاز: ${preferenceError instanceof Error ? preferenceError.message : "خطأ غير معروف."}`);
		}
	}

	function resetWidgetOrder(): void {
		const defaultOrder = [...DASHBOARD_WIDGET_IDS];
		setWidgetOrder(defaultOrder);
		try {
			window.localStorage.setItem(DASHBOARD_WIDGET_ORDER_STORAGE_KEY, JSON.stringify(defaultOrder));
			setError("");
		} catch (preferenceError) {
			setError(`تعذر حفظ ترتيب عناصر اللوحة على هذا الجهاز: ${preferenceError instanceof Error ? preferenceError.message : "خطأ غير معروف."}`);
		}
	}

	function updateWidgetOrder(sourceId: DashboardWidgetId, targetId: DashboardWidgetId): void {
		if (sourceId === targetId) return;
		const nextOrder = [...widgetOrder];
		const sourceIndex = nextOrder.indexOf(sourceId);
		const targetIndex = nextOrder.indexOf(targetId);
		if (sourceIndex < 0 || targetIndex < 0) return;
		nextOrder.splice(sourceIndex, 1);
		nextOrder.splice(targetIndex, 0, sourceId);
		setWidgetOrder(nextOrder);
		try {
			window.localStorage.setItem(DASHBOARD_WIDGET_ORDER_STORAGE_KEY, JSON.stringify(nextOrder));
			setError("");
		} catch (preferenceError) {
			setError(`تعذر حفظ ترتيب عناصر اللوحة على هذا الجهاز: ${preferenceError instanceof Error ? preferenceError.message : "خطأ غير معروف."}`);
		}
	}

	function widgetDragProps(widgetId: DashboardWidgetId) {
		return {
			"data-dashboard-widget": widgetId,
			draggable: true,
			style: { order: widgetOrder.indexOf(widgetId) },
			onDragStart: (event: DragEvent<HTMLElement>) => {
				setDraggedWidgetId(widgetId);
				event.dataTransfer.effectAllowed = "move";
				event.dataTransfer.setData("text/plain", widgetId);
			},
			onDragEnd: () => setDraggedWidgetId(null),
			onDragOver: (event: DragEvent<HTMLElement>) => {
				if (draggedWidgetId && draggedWidgetId !== widgetId) {
					event.preventDefault();
					event.dataTransfer.dropEffect = "move";
				}
			},
			onDrop: (event: DragEvent<HTMLElement>) => {
				event.preventDefault();
				const sourceId = event.dataTransfer.getData("text/plain") || draggedWidgetId;
				if (sourceId && isDashboardWidgetId(sourceId)) updateWidgetOrder(sourceId, widgetId);
				setDraggedWidgetId(null);
			},
		};
	}

	useEffect(() => {
		let isCurrent = true;
		void (async () => {
			try {
				const archiveList = await listArchives();
				if (!isCurrent) return;
				setArchives(archiveList);
				setArchiveStatus("connected");
				if (archiveList.length > 0) {
					const history = await Promise.all(
						archiveList.slice(0, TREND_SNAPSHOT_LIMIT).map((archive) => loadArchive(archive.id)),
					);
					if (!isCurrent) return;
					const latest = history[0];
					const previous = history[1] ?? null;
					setResult(archivedResult(latest));
					setLastUpdatedAt(latest.createdAt);
					setCurrentArchiveId(latest.id);
					setPreviousSnapshot(previous?.rows ?? null);
					setHistorySnapshots(history.map(({ id, createdAt, reportDate, rows }) => ({ id, createdAt, reportDate, rows })));
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
						setActiveOwnerId(OWNER_FILTERS[0].id);
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
						setActiveOwnerId(OWNER_FILTERS[0].id);
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

	const dateRangeArchives = useMemo(
		() => archives.filter((archive) =>
			archive.reportDate !== null &&
			archive.reportDate >= appliedDateRange.start &&
			archive.reportDate <= appliedDateRange.end,
		),
		[archives, appliedDateRange],
	);

	useEffect(() => {
		if (archives.length === 0) return;
		const requestId = ++dateRangeRequestId.current;
		let isCurrent = true;
		setIsDateRangeLoading(true);
		setResult(null);
		setCurrentArchiveId(null);
		setLastUpdatedAt(null);
		setPreviousSnapshot(null);
		setHistorySnapshots([]);
		setComparison(null);
		setSelectedArchiveIds(["", ""]);

		if (dateRangeArchives.length === 0) {
			setIsDateRangeLoading(false);
			return () => {
				isCurrent = false;
			};
		}

		void (async () => {
			try {
				const latestArchive = dateRangeArchives[0];
				const latestArchiveIndex = archives.findIndex((archive) => archive.id === latestArchive.id);
				const previousArchive = archives[latestArchiveIndex + 1];
				const snapshotsToLoad = dateRangeArchives.map((archive) => loadArchive(archive.id));
				if (previousArchive && !dateRangeArchives.some((archive) => archive.id === previousArchive.id)) {
					snapshotsToLoad.push(loadArchive(previousArchive.id));
				}
				const loadedSnapshots = await Promise.all(snapshotsToLoad);
				if (!isCurrent || requestId !== dateRangeRequestId.current) return;
				const rangedSnapshots = loadedSnapshots
					.filter((snapshot) => dateRangeArchives.some((archive) => archive.id === snapshot.id))
					.sort(compareByBalanceDate);
				const latest = rangedSnapshots[0];
				const previous = loadedSnapshots.find((snapshot) => snapshot.id === previousArchive?.id) ?? null;
				if (!latest) throw new Error("لم يتم العثور على بيانات ضمن النطاق الزمني المختار.");
				setResult(archivedResult(latest));
				setCurrentArchiveId(latest.id);
				setLastUpdatedAt(latest.createdAt);
				setPreviousSnapshot(previous?.rows ?? null);
				setHistorySnapshots(rangedSnapshots.map(({ id, createdAt, reportDate, rows }) => ({
					id,
					createdAt,
					reportDate,
					rows,
				})));
				setSelectedArchiveIds([rangedSnapshots[1]?.id ?? "", latest.id]);
				setError("");
			} catch (rangeError) {
				if (isCurrent && requestId === dateRangeRequestId.current) {
					setError(`تعذر تحميل بيانات النطاق الزمني: ${rangeError instanceof Error ? rangeError.message : "خطأ غير معروف."}`);
				}
			} finally {
				if (isCurrent && requestId === dateRangeRequestId.current) setIsDateRangeLoading(false);
			}
		})();
		return () => {
			isCurrent = false;
		};
	}, [archives, appliedDateRange, dateRangeArchives]);

	useEffect(() => {
		if (archives.length > 0 || archiveStatus !== "local") return;
		let isCurrent = true;
		setIsDateRangeLoading(true);
		setComparison(null);
		void loadSavedDashboard()
			.then((saved) => {
				if (!isCurrent) return;
				if (saved?.result.reportDate &&
					saved.result.reportDate >= appliedDateRange.start &&
					saved.result.reportDate <= appliedDateRange.end) {
					setResult(saved.result);
					setLastUpdatedAt(saved.updatedAt);
					setHistorySnapshots([{
						id: currentArchiveId ?? "local",
						createdAt: saved.updatedAt,
						reportDate: saved.result.reportDate,
						rows: saved.result.rows,
					}]);
				} else {
					setResult(null);
					setLastUpdatedAt(null);
					setHistorySnapshots([]);
				}
				setError("");
			})
			.catch((storageError: unknown) => {
				if (isCurrent) {
					setResult(null);
					setError(`تعذر تحميل النسخة المحلية للنطاق الزمني: ${storageError instanceof Error ? storageError.message : "خطأ غير معروف."}`);
				}
			})
			.finally(() => {
				if (isCurrent) setIsDateRangeLoading(false);
			});
		return () => {
			isCurrent = false;
		};
	}, [archives.length, archiveStatus, appliedDateRange, currentArchiveId]);

	const selectedRows = useMemo(() => {
		if (!result) return [];
		return result.rows.filter((customer) => {
			const ownerId = ownerForSegment(customer.segment);
			return ownerId
				? selectedOwners.has(ownerId)
				: showOtherSegments;
		});
	}, [result, selectedOwners, showOtherSegments]);

	const dashboardResult = result;
	const dashboardRows = useMemo(() => dashboardResult?.rows ?? [], [dashboardResult]);
	const dashboardPreviousRows = comparison?.first.rows ?? previousSnapshot;

	const representativeMetrics = useMemo<RepresentativeMetrics[]>(() => OWNER_FILTERS.map((owner) => {
		const rows = dashboardRows.filter((customer) => ownerForSegment(customer.segment) === owner.id);
		const previousRows = (dashboardPreviousRows ?? []).filter((customer) => ownerForSegment(customer.segment) === owner.id);
		const total = totalBalance(rows);
		const previousTotal = dashboardPreviousRows ? totalBalance(previousRows) : null;
		const changeAmount = previousTotal === null ? null : total - previousTotal;
		return {
			id: owner.id,
			name: owner.name,
			rows,
			total,
			under21: rows.reduce((sum, customer) => sum + sumBucketsThrough(customer, 3), 0),
			under45: rows.reduce((sum, customer) => sum + sumBucketsThrough(customer, 6), 0),
			over45: rows.reduce((sum, customer) => sum + sumBucketsFrom(customer, OVER_45_BUCKET_START), 0),
			over45Accounts: rows.filter((customer) => sumBucketsFrom(customer, OVER_45_BUCKET_START) > 0).length,
			previousTotal,
			changeAmount,
			changePercent: previousTotal === null || previousTotal === 0
				? null
				: ((changeAmount ?? 0) / previousTotal) * 100,
		};
	}), [dashboardPreviousRows, dashboardRows]);

	const activeRepresentativeRows = rowsForRepresentativeView(
		dashboardRows,
		activeOwnerId,
		selectedOwners,
		showOtherSegments,
	);
	const activeRepresentativeName = activeOwnerId === "all"
		? selectedOwners.size === OWNER_FILTERS.length
			? showOtherSegments ? "المندوبون الستة وباقي الفروع" : "المندوبون الستة"
			: "المندوبون المحددون"
		: representativeMetrics.find((owner) => owner.id === activeOwnerId)?.name ?? "المندوب";
	const activeRepresentativeTotal = totalBalance(activeRepresentativeRows);
	const activeRepresentativeUnder21 = activeRepresentativeRows.reduce(
		(sum, customer) => sum + sumBucketsThrough(customer, 3),
		0,
	);
	const activeRepresentativeUnder30 = activeRepresentativeRows.reduce(
		(sum, customer) => sum + sumBucketsThrough(customer, 4),
		0,
	);
	const activeRepresentativeOver45 = activeRepresentativeRows.reduce(
		(sum, customer) => sum + sumBucketsFrom(customer, OVER_45_BUCKET_START),
		0,
	);
	const activeRepresentative30To45 = activeRepresentativeRows.reduce(
		(sum, customer) => sum + amountBetweenBuckets(customer, OVER_30_BUCKET_START, OVER_45_BUCKET_START - 1),
		0,
	);
	const previousRowsForActive = rowsForRepresentativeView(
		dashboardPreviousRows ?? [],
		activeOwnerId,
		selectedOwners,
		showOtherSegments,
	);
	const previousTotalForActive = dashboardPreviousRows ? totalBalance(previousRowsForActive) : null;
	const activeChangeAmount = previousTotalForActive === null
		? null
		: activeRepresentativeTotal - previousTotalForActive;
	const activeChangePercent = previousTotalForActive === null || previousTotalForActive === 0
		? null
		: ((activeChangeAmount ?? 0) / previousTotalForActive) * 100;
	const activeBucketTotals = (dashboardResult?.bucketNames ?? []).map((name, index) => ({
		name,
		amount: activeRepresentativeRows.reduce((sum, customer) => sum + (customer.buckets[index] ?? 0), 0),
	}));
	const hasVisibleKpis = (
		visibleWidgets["kpi-total"] ||
		visibleWidgets["kpi-under21"] ||
		visibleWidgets["kpi-under30"] ||
		visibleWidgets["kpi-30-45"] ||
		visibleWidgets["kpi-over45"] ||
		visibleWidgets["kpi-commission"]
	);
	const snapshotsForTrend = comparison
		? comparison.trendSnapshots
		: historySnapshots.length > 0
			? historySnapshots
			: result && lastUpdatedAt
				? [{ id: currentArchiveId ?? "local", createdAt: lastUpdatedAt, reportDate: result.reportDate ?? null, rows: result.rows }]
				: [];
	const activeTrend = snapshotsForTrend.map((snapshot) => ({
		id: snapshot.id,
		createdAt: snapshot.createdAt,
		reportDate: snapshot.reportDate,
		total: totalBalance(rowsForRepresentativeView(
			snapshot.rows,
			activeOwnerId,
			selectedOwners,
			showOtherSegments,
		)),
	})).sort((first, second) =>
		(first.reportDate ?? first.createdAt.slice(0, 10)).localeCompare(second.reportDate ?? second.createdAt.slice(0, 10)) ||
		first.createdAt.localeCompare(second.createdAt),
	);
	const trendMinimum = Math.min(0, ...activeTrend.map((snapshot) => snapshot.total));
	const trendMaximum = Math.max(1, ...activeTrend.map((snapshot) => snapshot.total));
	const trendAxisLabels = [trendMaximum, (trendMinimum + trendMaximum) / 2, trendMinimum];
	const trendPoints = activeTrend.map((snapshot, index) => {
		const x = activeTrend.length <= 1 ? 440 : 100 + (index / (activeTrend.length - 1)) * 680;
		const y = 170 - ((snapshot.total - trendMinimum) / (trendMaximum - trendMinimum)) * 145;
		return { ...snapshot, x, y };
	});
	const trendLabelInterval = Math.max(1, Math.ceil(trendPoints.length / 7));
	const trendLabelPoints = trendPoints.filter((_, index) =>
		index % trendLabelInterval === 0 || index === trendPoints.length - 1,
	);
	const representativeComparisonValues = comparisonOwnerIds.flatMap((ownerId) =>
		historySnapshots.map((snapshot) =>
			valueForRepresentativeMetric(snapshot.rows, ownerId, representativeComparisonMetric),
		),
	);
	const representativeComparisonMinimum = Math.min(0, ...representativeComparisonValues);
	const representativeComparisonMaximum = Math.max(1, ...representativeComparisonValues);
	const representativeComparisonSeries = comparisonOwnerIds.map((ownerId, seriesIndex) => {
		const owner = OWNER_FILTERS.find((candidate) => candidate.id === ownerId) ?? OWNER_FILTERS[seriesIndex];
		const points = historySnapshots.map((snapshot, index) => {
			const value = valueForRepresentativeMetric(snapshot.rows, ownerId, representativeComparisonMetric);
			const x = historySnapshots.length <= 1 ? 440 : 100 + (index / (historySnapshots.length - 1)) * 680;
			const y = 170 - ((value - representativeComparisonMinimum) /
				(representativeComparisonMaximum - representativeComparisonMinimum)) * 145;
			return { id: snapshot.id, reportDate: snapshot.reportDate, value, x, y };
		});
		return {
			id: ownerId,
			name: owner.name,
			color: REPRESENTATIVE_COMPARISON_COLORS[seriesIndex],
			points,
			path: points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`).join(" "),
		};
	});
	const representativeComparisonAxisLabels = [
		representativeComparisonMaximum,
		(representativeComparisonMinimum + representativeComparisonMaximum) / 2,
		representativeComparisonMinimum,
	];
	const representativeComparisonLabelInterval = Math.max(1, Math.ceil(historySnapshots.length / 7));
	const representativeComparisonDateLabels = historySnapshots.filter((_, index) =>
		index % representativeComparisonLabelInterval === 0 || index === historySnapshots.length - 1,
	);
	const trendHeadlineBalance = activeRepresentativeTotal;
	const trendPath = trendPoints.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`).join(" ");
	const topFiveForActive = activeRepresentativeRows
		.map((customer) => ({
			...customer,
			over45: sumBucketsFrom(customer, OVER_45_BUCKET_START),
			under45: sumBucketsThrough(customer, 6),
		}))
		.filter((customer) => customer.over45 > 0)
		.sort((first, second) => second.over45 - first.over45)
		.slice(0, 5);
	const activeOver45Total = activeRepresentativeOver45;
	const topFiveOver45 = topFiveForActive.reduce((sum, customer) => sum + customer.over45, 0);
	const previous30PlusByAccount = new Map(
		previousRowsForActive.map((customer) => [
			customer.account,
			sumBucketsFrom(customer, OVER_30_BUCKET_START),
		]),
	);
	const activeNewlyOver30Rows = activeRepresentativeRows.filter((customer) =>
		(previous30PlusByAccount.get(customer.account) ?? 0) <= 0 &&
		sumBucketsFrom(customer, OVER_30_BUCKET_START) > 0,
	);
	const activeNewlyOver30Balance = activeNewlyOver30Rows.reduce(
		(sum, customer) => sum + sumBucketsFrom(customer, OVER_30_BUCKET_START),
		0,
	);
	const positiveAgingBuckets = activeBucketTotals.filter((bucket) => bucket.amount > 0);
	const positiveAgingTotal = positiveAgingBuckets.reduce((sum, bucket) => sum + bucket.amount, 0);
	let agingAngle = 0;
	const agingDonut = positiveAgingTotal > 0
		? `conic-gradient(${positiveAgingBuckets.map((bucket) => {
			const startAngle = agingAngle;
			agingAngle += (bucket.amount / positiveAgingTotal) * 360;
			return `${AGING_CHART_COLORS[activeBucketTotals.indexOf(bucket)]} ${startAngle}deg ${agingAngle}deg`;
		}).join(", ")})`
		: "conic-gradient(#e9efea 0deg 360deg)";
	const maxCustomerBucket = Math.max(
		1,
		...activeRepresentativeRows.flatMap((customer) => customer.buckets.map((amount) => Math.abs(amount))),
	);
	const maxBucket = Math.max(0, ...activeBucketTotals.map((bucket) => bucket.amount));
	const activeHeatmapRows = [...activeRepresentativeRows].sort((first, second) => {
		let difference: number;
		switch (heatmapSort.column) {
			case "customer":
				difference = (first.name || first.account).localeCompare(second.name || second.account, "ar");
				break;
			case "segment":
				difference = first.segment.localeCompare(second.segment, "ar");
				break;
			case "bucket":
				difference = (first.buckets[heatmapSort.bucketIndex ?? 0] ?? 0) -
					(second.buckets[heatmapSort.bucketIndex ?? 0] ?? 0);
				break;
			case "total":
				difference = first.total - second.total;
				break;
		}
		return heatmapSort.direction === "desc" ? -difference : difference;
	});

	const activeSegments = useMemo(() => {
		const totalsBySegment = new Map<string, number>();
		for (const customer of activeRepresentativeRows) {
			totalsBySegment.set(
				customer.segment,
				(totalsBySegment.get(customer.segment) ?? 0) + customer.total,
			);
		}
		return [...totalsBySegment.entries()].sort((a, b) => b[1] - a[1]);
	}, [activeRepresentativeRows]);

	const selectedTotal = useMemo(
		() => selectedRows.reduce((sum, customer) => sum + customer.total, 0),
		[selectedRows],
	);
	function selectRepresentative(ownerId: string): void {
		setActiveOwnerId(ownerId);
		if (ownerId === "all") {
			setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
			setShowOtherSegments(false);
		} else {
			setSelectedOwners(new Set([ownerId]));
			setShowOtherSegments(false);
		}
	}
	const maxSegment = Math.max(0, ...activeSegments.slice(0, 8).map(([, amount]) => amount));
	const otherSegments = activeSegments.slice(8);
	const otherSegmentTotal = otherSegments.reduce((sum, [, amount]) => sum + amount, 0);
	const formattedBalanceDate = result
		? formatReportDate(result.reportDate)
		: "لا توجد بيانات محفوظة";

	async function handleFileChange(key: UploadKey, event: ChangeEvent<HTMLInputElement>) {
		const file = event.currentTarget.files?.[0] ?? null;
		event.currentTarget.value = "";
		const nextUploads = { ...uploads, [key]: file };
		setUploads(nextUploads);
		setError("");
		setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
		setShowOtherSegments(false);
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
				const processedBalanceDate = processed.reportDate ?? updatedAt.slice(0, 10);
				const currentBalanceDate = result?.reportDate || lastUpdatedAt?.slice(0, 10) || "";
				if (!result || processedBalanceDate >= currentBalanceDate) {
					setResult(processed);
					setLastUpdatedAt(updatedAt);
				}
				setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
				setShowOtherSegments(false);
				setComparison(null);
				try {
					const archive = await saveArchive(processed);
					const nextArchives = [archive, ...archives.filter((item) => item.id !== archive.id)]
						.sort(compareByBalanceDate);
					const latestPair = await loadLatestArchivePair(nextArchives);
					if (!latestPair) throw new Error("تعذر تحديد أحدث نسخة بعد الحفظ.");
					const { latest, previous } = latestPair;
					const latestResult = archivedResult(latest);
					setResult(latestResult);
					setPreviousSnapshot(previous?.rows ?? null);
					setCurrentArchiveId(latest.id);
					setLastUpdatedAt(latest.createdAt);
					setArchives(nextArchives);
					setHistorySnapshots((current) => [
						...(previous ? [{
							id: previous.id,
							createdAt: previous.createdAt,
							reportDate: previous.reportDate,
							rows: previous.rows,
						}] : []),
						{
							id: latest.id,
							createdAt: latest.createdAt,
							reportDate: latest.reportDate,
							rows: latest.rows,
						},
						...current.filter((snapshot) => snapshot.id !== latest.id && snapshot.id !== previous?.id),
					].sort(compareByBalanceDate).slice(0, TREND_SNAPSHOT_LIMIT));
					setSelectedArchiveIds([previous?.id ?? "", latest.id]);
					await saveDashboard({ result: latestResult, updatedAt: latest.createdAt });
					setArchiveStatus("connected");
					setError("");
				} catch (storageError) {
					setArchiveStatus("local");
					let localSaveMessage = "";
					if (!result || processedBalanceDate >= currentBalanceDate) {
						try {
							await saveDashboard({ result: processed, updatedAt });
						} catch (localError) {
							localSaveMessage = ` وتعذر أيضاً حفظ النسخة المحلية: ${localError instanceof Error ? localError.message : "خطأ غير معروف."}`;
						}
					}
					setError(`تمت المعالجة محلياً لكن تعذر حفظ أحدث نسخة في Cloudflare: ${storageError instanceof Error ? storageError.message : "خطأ غير معروف."}${localSaveMessage}`);
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
			const nextArchives = [archive, ...archives.filter((item) => item.id !== archive.id)]
				.sort(compareByBalanceDate);
			const latestPair = await loadLatestArchivePair(nextArchives);
			if (!latestPair) throw new Error("تعذر تحديد أحدث نسخة بعد النشر.");
			const { latest, previous } = latestPair;
			setArchives(nextArchives);
			const latestResult = archivedResult(latest);
			setResult(latestResult);
			setPreviousSnapshot(previous?.rows ?? null);
			setCurrentArchiveId(latest.id);
			setLastUpdatedAt(latest.createdAt);
			setSelectedArchiveIds([previous?.id ?? "", latest.id]);
			setHistorySnapshots((current) => [
				...(previous ? [{
					id: previous.id,
					createdAt: previous.createdAt,
					reportDate: previous.reportDate,
					rows: previous.rows,
				}] : []),
				{
					id: latest.id,
					createdAt: latest.createdAt,
					reportDate: latest.reportDate,
					rows: latest.rows,
				},
				...current.filter((snapshot) => snapshot.id !== latest.id && snapshot.id !== previous?.id),
			].sort(compareByBalanceDate).slice(0, TREND_SNAPSHOT_LIMIT));
			setComparison(null);
			await saveDashboard({ result: latestResult, updatedAt: latest.createdAt });
			setArchiveStatus("connected");
			setError("");
		} catch (archiveError) {
			setError(`تعذر نشر النسخة المحلية إلى الأرشيف: ${archiveError instanceof Error ? archiveError.message : "خطأ غير معروف."}`);
		}
	}

	async function removeArchive(archive: ArchiveSummary): Promise<void> {
		if (!window.confirm(`هل تريد حذف نسخة ${formatReportDate(archive.reportDate)} نهائياً من الأرشيف؟`)) return;
		try {
			await deleteArchive(archive.id);
			const remaining = archives.filter((item) => item.id !== archive.id);
			setArchives(remaining);
			setHistorySnapshots((current) => current.filter((snapshot) => snapshot.id !== archive.id));
			setComparison(null);
			const latestPair = await loadLatestArchivePair(remaining);
			if (latestPair) {
				const { latest, previous } = latestPair;
				const latestResult = archivedResult(latest);
				setResult(latestResult);
				setCurrentArchiveId(latest.id);
				setLastUpdatedAt(latest.createdAt);
				setPreviousSnapshot(previous?.rows ?? null);
				setSelectedArchiveIds([previous?.id ?? "", latest.id]);
				await saveDashboard({ result: latestResult, updatedAt: latest.createdAt });
			} else {
				setResult(null);
				setCurrentArchiveId(null);
				setLastUpdatedAt(null);
				setPreviousSnapshot(null);
				setSelectedArchiveIds(["", ""]);
			}
			setError("");
		} catch (archiveError) {
			setError(`تعذر حذف النسخة: ${archiveError instanceof Error ? archiveError.message : "خطأ غير معروف."}`);
		}
	}

	async function compareArchives(): Promise<void> {
		const firstId = selectedArchiveIds[0];
		const latestId = dateRangeArchives[0]?.id;
		const baselineIndex = dateRangeArchives.findIndex((archive) => archive.id === firstId);
		if (!firstId || !latestId || baselineIndex < 1) {
			setError("اختر فترة أساس ضمن النطاق وغير الأحدث للمقارنة.");
			return;
		}
		setIsComparing(true);
		try {
			const datasets = await Promise.all(
				dateRangeArchives.slice(0, baselineIndex + 1).map((archive) => loadArchive(archive.id)),
			);
			const first = datasets.find((dataset) => dataset.id === firstId);
			const second = datasets.find((dataset) => dataset.id === latestId);
			if (!first || !second) throw new Error("تعذر تحميل فترة الأساس أو أحدث فترة من الأرشيف.");
			const trendSnapshots = datasets.map(({ id, createdAt, reportDate, rows }) => ({
				id,
				createdAt,
				reportDate,
				rows,
			}));
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
			setComparison({ first, second, trendSnapshots, rows });
			setError("");
		} catch (archiveError) {
			setError(`تعذرت مقارنة النسختين: ${archiveError instanceof Error ? archiveError.message : "خطأ غير معروف."}`);
		} finally {
			setIsComparing(false);
		}
	}

	function toggleHeatmapSort(column: HeatmapSort["column"], bucketIndex?: number): void {
		setHeatmapSort((current) => {
			const sameColumn = current.column === column && current.bucketIndex === bucketIndex;
			return {
				column,
				bucketIndex,
				direction: sameColumn && current.direction === "desc" ? "asc" : "desc",
			};
		});
	}

	function heatmapSortIndicator(column: HeatmapSort["column"], bucketIndex?: number): string {
		return heatmapSort.column === column && heatmapSort.bucketIndex === bucketIndex
			? heatmapSort.direction === "desc" ? " ↓" : " ↑"
			: "";
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
					<span>تاريخ الرصيد</span>
					<strong>{isRestoring ? "جاري الاستعادة…" : formattedBalanceDate}</strong>
				</div>
			</header>
			<nav className="app-tabs" role="tablist" aria-label="أقسام التطبيق">
				<button
					id="dashboard-tab"
					className={`app-tab${activeTab === "dashboard" ? " active" : ""}`}
					type="button"
					role="tab"
					aria-selected={activeTab === "dashboard"}
					aria-controls="dashboard-tab-panel"
					onClick={() => setActiveTab("dashboard")}
				>
					لوحة تقييم الأداء
				</button>
				<button
					id="daily-sales-tab"
					className={`app-tab${activeTab === "daily-sales" ? " active" : ""}`}
					type="button"
					role="tab"
					aria-selected={activeTab === "daily-sales"}
					aria-controls="daily-sales-tab-panel"
					onClick={() => setActiveTab("daily-sales")}
				>
					متابعة المبيعات اليومية
				</button>
			</nav>
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

			<div id="dashboard-tab-panel" role="tabpanel" aria-labelledby="dashboard-tab" hidden={activeTab !== "dashboard"}>
			<section className="sales-dashboard-hero" id="top" aria-labelledby="sales-dashboard-title">
				<div className="sales-dashboard-heading">
					<div className="sales-heading-title">
						<h1 id="sales-dashboard-title">لوحة تقييم الأداء</h1>
						<div className="new-over30-badge" title={dashboardPreviousRows
							? `إجمالي رصيد +30 لهذه الحسابات: ${formatAmount(activeNewlyOver30Balance)}`
							: "اختر فترة مقارنة لعرض الحسابات الجديدة فوق 30 يوماً"}>
							<span>دخلت +30</span>
							<strong>{dashboardPreviousRows ? activeNewlyOver30Rows.length : "—"}</strong>
							<small>حساب</small>
						</div>
					</div>
					<div className="dashboard-toolbar">
						<div className="date-range-controls">
							<label>
								<span>نطاق التاريخ</span>
								<select
									aria-label="النطاق الزمني للوحة"
									value={dateRangePreset}
									onChange={(event) => {
										const preset = event.currentTarget.value;
										if (!isDateRangePreset(preset)) return;
										setDateRangePreset(preset);
										if (preset !== "custom") {
											const range = getDateRange(preset);
											setCustomDateRange(range);
											setAppliedDateRange(range);
										}
									}}
								>
									<option value="today">اليوم</option>
									<option value="yesterday">الأمس</option>
									<option value="last-7-days">آخر 7 أيام</option>
									<option value="last-30-days">آخر 30 يوماً</option>
									<option value="last-12-months">آخر 12 شهراً</option>
									<option value="custom">نطاق مخصص</option>
								</select>
							</label>
							{dateRangePreset === "custom" && (
								<>
									<label>
										<span>من</span>
										<input
											aria-label="بداية النطاق المخصص"
											type="date"
											value={customDateRange.start}
											max={customDateRange.end}
											onChange={(event) => {
												const start = event.currentTarget.value;
												setCustomDateRange((current) => ({ ...current, start }));
											}}
										/>
									</label>
									<label>
										<span>إلى</span>
										<input
											aria-label="نهاية النطاق المخصص"
											type="date"
											value={customDateRange.end}
											min={customDateRange.start}
											onChange={(event) => {
												const end = event.currentTarget.value;
												setCustomDateRange((current) => ({ ...current, end }));
											}}
										/>
									</label>
									<button
										className="button button-dark"
										type="button"
										disabled={!customDateRange.start || !customDateRange.end || customDateRange.start > customDateRange.end}
										onClick={() => {
											setAppliedDateRange(customDateRange);
											setError("");
										}}
									>
										تطبيق النطاق
									</button>
								</>
							)}
						</div>
						{dateRangeArchives.length > 1 && (
							<div className="comparison-controls dashboard-comparison-controls">
								<span className="comparison-strip-label">مقارنة أحدث تاريخ مع:</span>
								<label>
									<span>فترة الأساس</span>
									<select
										aria-label="فترة الأساس للمقارنة"
										value={selectedArchiveIds[0]}
										onChange={(event) => {
											setSelectedArchiveIds([event.currentTarget.value, dateRangeArchives[0].id]);
											setComparison(null);
										}}
									>
										<option value="">اختر فترة</option>
										{dateRangeArchives.slice(1).map((archive) => (
											<option key={archive.id} value={archive.id}>
												{archiveOptionLabel(archive)} · {formatAmount(archive.total)}
											</option>
										))}
									</select>
								</label>
								<span className="comparison-strip-label comparison-latest-label">
									الأحدث: {archiveOptionLabel(dateRangeArchives[0])}
								</span>
								<button
									className="button button-dark"
									type="button"
									disabled={isComparing || !selectedArchiveIds[0] || selectedArchiveIds[0] === dateRangeArchives[0].id}
									onClick={() => void compareArchives()}
								>
									{isComparing ? "جارٍ…" : "تطبيق المقارنة"}
								</button>
							</div>
						)}
						<div className="dashboard-widget-customizer">
							<button
								className="button button-quiet dashboard-widget-toggle"
								type="button"
								aria-expanded={isWidgetSettingsOpen}
								aria-controls="dashboard-widget-settings"
								onClick={() => setIsWidgetSettingsOpen((open) => !open)}
							>
								{isWidgetSettingsOpen ? "إغلاق تخصيص اللوحة" : "＋ تخصيص اللوحة"}
							</button>
							{isWidgetSettingsOpen && (
								<div className="dashboard-widget-settings" id="dashboard-widget-settings">
									<div className="dashboard-widget-settings-heading">
										<div>
											<strong>عناصر اللوحة</strong>
											<p>أظهر أو أخفِ العناصر، واسحب أي بطاقة أو رسم إلى موضعه الجديد. تُحفظ التغييرات على هذا الجهاز.</p>
										</div>
											<div className="dashboard-widget-reset-actions">
												<button className="button button-quiet" type="button" onClick={resetWidgetVisibility}>
													إظهار الكل
												</button>
												<button className="button button-quiet" type="button" onClick={resetWidgetOrder}>
													استعادة الترتيب
												</button>
											</div>
									</div>
									<div className="dashboard-widget-groups">
										{DASHBOARD_WIDGET_GROUPS.map((group) => (
											<fieldset key={group.title}>
												<legend>{group.title}</legend>
												{group.items.map(({ id, label }) => (
													<label key={id}>
														<input
															type="checkbox"
															checked={visibleWidgets[id]}
															onChange={(event) => updateWidgetVisibility(id, event.currentTarget.checked)}
														/>
														<span>{label}</span>
													</label>
												))}
											</fieldset>
										))}
									</div>
								</div>
							)}
						</div>
					</div>
				</div>
				{isDateRangeLoading && <p className="date-range-status" role="status">جارٍ تحميل بيانات النطاق المحدد…</p>}
				{!isDateRangeLoading && (archives.length > 0 || archiveStatus === "local") && dateRangeArchives.length === 0 && !result && (
					<p className="date-range-status" role="status">لا توجد نسخة بتاريخ رصيد ضمن النطاق المحدد.</p>
				)}
				{result && <div className="representative-selector" role="group" aria-label="اختيار مندوب المبيعات">
					<button
						type="button"
						className={`representative-card representative-all-card${activeOwnerId === "all" && !showOtherSegments && selectedOwners.size === OWNER_FILTERS.length ? " active" : ""}`}
						aria-pressed={activeOwnerId === "all" && !showOtherSegments && selectedOwners.size === OWNER_FILTERS.length}
						onClick={() => selectRepresentative("all")}
					>
						<span>كل المندوبين</span>
						<strong>{formatAmount(totalBalance(rowsForRepresentativeView(dashboardRows, "all", selectedOwners, showOtherSegments)))}</strong>
						<small>{rowsForRepresentativeView(dashboardRows, "all", selectedOwners, showOtherSegments).length} حساب</small>
					</button>
					{representativeMetrics.map((owner) => (
						<button
							type="button"
							className={`representative-card${activeOwnerId === owner.id ? " active" : ""}`}
							aria-pressed={activeOwnerId === owner.id}
							key={owner.id}
							onClick={() => selectRepresentative(owner.id)}
						>
							<span>{owner.name}</span>
							<strong>{formatAmount(owner.total)}</strong>
							<small className={owner.changeAmount === null || owner.changeAmount === 0 ? "" : owner.changeAmount > 0 ? "change-up" : "change-down"}>
								{owner.changePercent === null
									? owner.previousTotal === null ? "لا يوجد تحديث سابق" : owner.total > 0 ? "محفظة جديدة" : "لا تغيير"
									: `${owner.changeAmount !== null && owner.changeAmount > 0 ? "↑ " : owner.changeAmount !== null && owner.changeAmount < 0 ? "↓ " : ""}${owner.changePercent > 0 ? "+" : ""}${percentFormat.format(owner.changePercent)}٪ عن السابق`}
							</small>
						</button>
					))}
					<button
						type="button"
						className={`representative-card representative-other-card${showOtherSegments ? " active" : ""}`}
						aria-pressed={showOtherSegments}
						onClick={() => {
							setActiveOwnerId("all");
							setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
							setShowOtherSegments((visible) => !visible);
						}}
					>
						<span>باقي الفروع</span>
						<strong>{formatAmount(totalBalance(dashboardRows.filter((customer) => ownerForSegment(customer.segment) === null)))}</strong>
						<small>{dashboardRows.filter((customer) => ownerForSegment(customer.segment) === null).length} حساب</small>
					</button>
				</div>}
				{result ? (
					<div
						key={`${activeOwnerId}-${showOtherSegments}-${[...selectedOwners].sort().join(",")}`}
						className="representative-view-transition"
					>
						<div className={`rep-kpi-grid${hasVisibleKpis ? "" : " dashboard-widget-hidden"}`}>
							<article {...widgetDragProps("kpi-total")} className={`rep-kpi-card dashboard-widget-draggable${visibleWidgets["kpi-total"] ? "" : " dashboard-widget-hidden"}`}>
								<span>إجمالي الرصيد</span>
								<strong title={comparison ? `رصيد فترة العرض ${formatReportDate(comparison.second.reportDate)}` : "الرصيد الحالي"}>
									{formatAmount(trendHeadlineBalance)}
								</strong>
								<small className={`risk-change${activeChangeAmount === null || activeChangeAmount === 0 ? "" : activeChangeAmount > 0 ? " change-up" : " change-down"}`}>
									{activeChangeAmount === null
										? dashboardPreviousRows ? "لا توجد قيمة سابقة" : "لا توجد مقارنة بعد"
										: `${activeChangeAmount > 0 ? "↑ " : activeChangeAmount < 0 ? "↓ " : ""}${activeChangeAmount > 0 ? "+" : ""}${formatAmount(activeChangeAmount)} (${activeChangePercent === null ? "—" : `${activeChangePercent > 0 ? "+" : ""}${percentFormat.format(activeChangePercent)}٪`}) ${activeChangeAmount > 0 ? "· ارتفاع يحتاج متابعة" : activeChangeAmount < 0 ? "· انخفاض محتمل" : "· دون تغيير"}`}
								</small>
								<small className="rep-kpi-footnote">تغير الرصيد لا يثبت التحصيل وحده</small>
							</article>
							<article {...widgetDragProps("kpi-under21")} className={`rep-kpi-card rep-kpi-under21 dashboard-widget-draggable${visibleWidgets["kpi-under21"] ? "" : " dashboard-widget-hidden"}`}>
								<span>الرصيد حتى 21 يوماً</span>
								<strong>{formatAmount(activeRepresentativeUnder21)}</strong>
								<small>{formatPercent(activeRepresentativeUnder21, activeRepresentativeTotal)} من المحفظة · Current وحتى فترة أقل من 21</small>
							</article>
							<article {...widgetDragProps("kpi-under30")} className={`rep-kpi-card rep-kpi-under45 dashboard-widget-draggable${visibleWidgets["kpi-under30"] ? "" : " dashboard-widget-hidden"}`}>
								<span>الرصيد حتى 30 يوماً</span>
								<strong>{formatAmount(activeRepresentativeUnder30)}</strong>
								<small>{formatPercent(activeRepresentativeUnder30, activeRepresentativeTotal)} من المحفظة · مجموع الفترات حتى 30 يوماً</small>
							</article>
							<article {...widgetDragProps("kpi-30-45")} className={`rep-kpi-card rep-kpi-warning dashboard-widget-draggable${visibleWidgets["kpi-30-45"] ? "" : " dashboard-widget-hidden"}`}>
								<span>الرصيد بين 30–45 يوماً</span>
								<strong>{formatAmount(activeRepresentative30To45)}</strong>
								<small>{formatPercent(activeRepresentative30To45, activeRepresentativeTotal)} من المحفظة · فترتا 30–37 و37–45 فقط</small>
								<small className="rep-kpi-commission">
									العمولة المعرّضة للفقد (1٪): <strong>{formatAmount(activeRepresentative30To45 * COMMISSION_RATE)}</strong>
								</small>
							</article>
							<article {...widgetDragProps("kpi-over45")} className={`rep-kpi-card rep-kpi-risk dashboard-widget-draggable${visibleWidgets["kpi-over45"] ? "" : " dashboard-widget-hidden"}`}>
								<span>الرصيد فوق 45 يوماً</span>
								<strong>{formatAmount(activeRepresentativeOver45)}</strong>
								<small>{formatPercent(activeRepresentativeOver45, activeRepresentativeTotal)} من المحفظة · {activeRepresentativeRows.filter((customer) => sumBucketsFrom(customer, OVER_45_BUCKET_START) > 0).length} حساب</small>
							</article>
							<article {...widgetDragProps("kpi-commission")} className={`rep-kpi-card rep-kpi-risk dashboard-widget-draggable${visibleWidgets["kpi-commission"] ? "" : " dashboard-widget-hidden"}`}>
								<span>العمولة المفقودة التقديرية · +45 · 1٪</span>
								<strong>{formatAmount(activeRepresentativeOver45 * COMMISSION_RATE)}</strong>
								<small>تقدير العمولة المرتبطة بالرصيد فوق 45 يوماً</small>
							</article>
						</div>
						<div className={`rep-chart-grid${visibleWidgets.trend || visibleWidgets["aging-pie"] ? "" : " dashboard-widget-hidden"}`}>
							<article {...widgetDragProps("trend")} className={`rep-chart-card dashboard-widget-draggable${visibleWidgets.trend ? "" : " dashboard-widget-hidden"}`}>
								<div className="rep-chart-heading">
									<div>
										<h3>اتجاه إجمالي الرصيد</h3>
										<p>{comparison ? `من ${formatReportDate(comparison.first.reportDate)} حتى أحدث فترة، مع عرض كل الفترات بينهما` : `حسب آخر ${activeTrend.length} لقطات محفوظة`} · لا يمثل التحصيل وحده</p>
									</div>
									<strong title={comparison ? `رصيد فترة العرض ${formatReportDate(comparison.second.reportDate)}` : "الرصيد الحالي"}>
										{formatAmount(trendHeadlineBalance)}
									</strong>
								</div>
								{activeTrend.length > 0 ? (
									<>
										<svg className="rep-line-chart" viewBox="0 0 780 190" role="img" aria-label={`اتجاه الرصيد للمندوب ${activeRepresentativeName}`}>
											{trendAxisLabels.map((amount, index) => {
												const y = [25, 96, 170][index];
												return (
													<g key={index}>
														<text className="rep-trend-axis-label" x="0" y={y + 4}>
															<title>{formatAmount(amount)}</title>
															{trendAxisAmountFormat.format(amount)}
														</text>
														<line x1="88" y1={y} x2="780" y2={y} />
													</g>
												);
											})}
											{activeTrend.length > 1 && <path className="rep-trend-path" d={trendPath} />}
											{trendPoints.map((point) => (
												<circle key={point.id} cx={point.x} cy={point.y} r="5">
													<title>{formatReportDate(point.reportDate)}: {formatAmount(point.total)}</title>
												</circle>
											))}
										</svg>
										<div className="rep-chart-labels">
											{trendLabelPoints.map((point, index) => (
												<span
													key={point.id}
													style={{ left: `${(point.x / 780) * 100}%` }}
													className={index === 0 ? "first" : index === trendLabelPoints.length - 1 ? "last" : ""}
													title={formatReportDate(point.reportDate)}
												>
													{formatTrendDate(point.reportDate)}
												</span>
											))}
										</div>
									</>
								) : <p className="rep-chart-empty">لا توجد بيانات تاريخية لهذا المندوب.</p>}
							</article>
							<article {...widgetDragProps("aging-pie")} className={`rep-chart-card dashboard-widget-draggable${visibleWidgets["aging-pie"] ? "" : " dashboard-widget-hidden"}`}>
								<div className="rep-chart-heading">
									<div>
										<h3>حصة الفترات من رصيد المندوب</h3>
										<p>توزيع نسبي حسب شرائح الأعمار · القيم السالبة مستبعدة من الدائرة</p>
									</div>
								</div>
								<div className="aging-donut-layout">
									<div
										className="aging-donut"
										style={{ background: agingDonut }}
										role="img"
										aria-label={`توزيع ${formatAmount(positiveAgingTotal)} على فترات أعمار الرصيد`}
									>
										<div className="aging-donut-center">
											<strong>{formatAmount(positiveAgingTotal)}</strong>
											<span>إجمالي موجب</span>
										</div>
									</div>
									<div className="aging-donut-legend">
										{activeBucketTotals.map((bucket, index) => (
											<div className="aging-legend-item" key={`${index}-${bucket.name}`}>
												<span className="aging-legend-dot" style={{ backgroundColor: AGING_CHART_COLORS[index % AGING_CHART_COLORS.length] }} />
												<span className="aging-legend-name">{bucket.name}</span>
												<strong>{formatPercent(Math.max(0, bucket.amount), positiveAgingTotal)}</strong>
												<small>{formatAmount(bucket.amount)}</small>
											</div>
										))}
									</div>
								</div>
							</article>
						</div>
						<article {...widgetDragProps("owner-comparison")} className={`rep-chart-card rep-owner-comparison-card dashboard-widget-draggable${visibleWidgets["owner-comparison"] ? "" : " dashboard-widget-hidden"}`}>
							<div className="rep-chart-heading rep-owner-comparison-heading">
								<div>
									<h3>مقارنة أداء مندوبين</h3>
									<p>مقارنة الاتجاه عبر تواريخ الرصيد ضمن النطاق المختار</p>
								</div>
								<div className="rep-owner-comparison-controls">
									<label>
										<span>المؤشر</span>
										<select
											aria-label="مؤشر مقارنة المندوبين"
											value={representativeComparisonMetric}
											onChange={(event) => {
												const metric = event.currentTarget.value;
												if (isRepresentativeComparisonMetric(metric)) {
													setRepresentativeComparisonMetric(metric);
												}
											}}
										>
											<option value="balance">إجمالي الرصيد</option>
											<option value="over30">الرصيد +30</option>
											<option value="over45">الرصيد +45</option>
											<option value="commission">عمولة +45 التقديرية</option>
										</select>
									</label>
									<label>
										<span>المندوب الأول</span>
										<select
											aria-label="المندوب الأول للمقارنة"
											value={comparisonOwnerIds[0]}
											onChange={(event) => {
												const ownerId = event.currentTarget.value;
												setComparisonOwnerIds((current) => ownerId === current[1]
													? [current[1], current[0]]
													: [ownerId, current[1]]);
											}}
										>
											{OWNER_FILTERS.map((owner) => (
												<option key={owner.id} value={owner.id}>{owner.name}</option>
											))}
										</select>
									</label>
									<label>
										<span>المندوب الثاني</span>
										<select
											aria-label="المندوب الثاني للمقارنة"
											value={comparisonOwnerIds[1]}
											onChange={(event) => {
												const ownerId = event.currentTarget.value;
												setComparisonOwnerIds((current) => ownerId === current[0]
													? [current[1], current[0]]
													: [current[0], ownerId]);
											}}
										>
											{OWNER_FILTERS.map((owner) => (
												<option key={owner.id} value={owner.id}>{owner.name}</option>
											))}
										</select>
									</label>
								</div>
							</div>
							<div className="rep-owner-comparison-legend">
								{representativeComparisonSeries.map((series) => (
									<span key={series.id}>
										<i style={{ backgroundColor: series.color }} />
										{series.name}
											<strong>{formatAmount(series.points[series.points.length - 1]?.value ?? 0)}</strong>
									</span>
								))}
							</div>
							{historySnapshots.length > 0 ? (
								<>
									<svg className="rep-line-chart" viewBox="0 0 780 190" role="img" aria-label="مقارنة أداء المندوبين عبر الزمن">
										{representativeComparisonAxisLabels.map((amount, index) => {
											const y = [25, 96, 170][index];
											return (
												<g key={index}>
													<text className="rep-trend-axis-label" x="0" y={y + 4}>
														<title>{formatAmount(amount)}</title>
														{trendAxisAmountFormat.format(amount)}
													</text>
													<line x1="88" y1={y} x2="780" y2={y} />
												</g>
											);
										})}
										{representativeComparisonSeries.map((series) => (
											<g key={series.id}>
												{series.points.length > 1 && (
													<path className="rep-trend-path" style={{ stroke: series.color }} d={series.path} />
												)}
												{series.points.map((point) => (
													<circle key={point.id} cx={point.x} cy={point.y} r="4" style={{ fill: series.color, stroke: series.color }}>
														<title>{series.name} · {formatReportDate(point.reportDate)}: {formatAmount(point.value)}</title>
													</circle>
												))}
											</g>
										))}
									</svg>
									<div className="rep-chart-labels">
										{representativeComparisonDateLabels.map((snapshot, index) => {
											const pointIndex = historySnapshots.findIndex((item) => item.id === snapshot.id);
											const x = historySnapshots.length <= 1
												? 440
												: 100 + (pointIndex / (historySnapshots.length - 1)) * 680;
											return (
												<span
													key={snapshot.id}
													style={{ left: `${(x / 780) * 100}%` }}
													className={index === 0 ? "first" : index === representativeComparisonDateLabels.length - 1 ? "last" : ""}
													title={formatReportDate(snapshot.reportDate)}
												>
													{formatTrendDate(snapshot.reportDate)}
												</span>
											);
										})}
									</div>
								</>
							) : <p className="rep-chart-empty">لا توجد لقطات ضمن النطاق الزمني لعرض المقارنة.</p>}
						</article>
						<div className={`rep-bottom-grid${visibleWidgets["balance-distribution"] || visibleWidgets["over45-concentration"] || visibleWidgets["owner-balances"] ? "" : " dashboard-widget-hidden"}`}>
							<article {...widgetDragProps("balance-distribution")} className={`rep-chart-card panel age-panel dashboard-widget-draggable${visibleWidgets["balance-distribution"] ? "" : " dashboard-widget-hidden"}`}>
								<div className="panel-heading">
									<div>
										<p className="panel-kicker">مجموع الأرصدة حسب الفترات</p>
										<h3>توزيع إجمالي الرصيد</h3>
									</div>
									<span className="panel-icon">▥</span>
								</div>
								<div className="age-chart">
									{activeBucketTotals.map((bucket, index) => (
										<div className="age-row" key={`${index}-${bucket.name}`}>
											<div className="age-label" title={bucket.name}>
												<span>{bucket.name}</span>
												<strong>{formatAmount(bucket.amount)}</strong>
											</div>
											<div className="bar-track">
												<div
													className={`bar-fill age-fill segment-color-${index % 5}${index >= OVER_30_BUCKET_START && index < OVER_45_BUCKET_START ? ` aging-warning-${index}` : ""}`}
													style={{ width: `${Math.max(0, (bucket.amount / Math.max(maxBucket, 1)) * 100)}%` }}
												/>
											</div>
										</div>
									))}
								</div>
								<div className="chart-footnote">
									<span>مجموع الفترات</span>
									<strong>{formatAmount(activeRepresentativeTotal)}</strong>
								</div>
							</article>
							<div className={`rep-side-stack${visibleWidgets["over45-concentration"] || visibleWidgets["owner-balances"] ? "" : " dashboard-widget-hidden"}`}>
								<article {...widgetDragProps("over45-concentration")} className={`rep-chart-card concentration-card dashboard-widget-draggable${visibleWidgets["over45-concentration"] ? "" : " dashboard-widget-hidden"}`}>
								<div className="rep-chart-heading">
									<div>
										<h3>تركيز متأخرات +45</h3>
										<p>أكبر الحسابات المتأخرة من رصيد المندوب</p>
									</div>
									<strong>{formatPercent(topFiveOver45, activeOver45Total)}</strong>
								</div>
								{topFiveForActive.length > 0 ? topFiveForActive.map((customer, index) => (
									<div className="concentration-row" key={customer.account}>
										<div>
											<span className="concentration-rank">{index + 1}</span>
											<span className="concentration-name">{customer.name || customer.account}</span>
											<strong>{formatAmount(customer.over45)}</strong>
										</div>
										<div className="bar-track">
											<div className="bar-fill priority-concentration-fill" style={{ width: `${Math.max(0, (customer.over45 / Math.max(activeOver45Total, 1)) * 100)}%` }} />
										</div>
									</div>
								)) : <p className="rep-chart-empty">لا توجد متأخرات +45 لدى هذا المندوب.</p>}
								<p className="rep-chart-footnote">أعلى 5 حسابات تشكل {formatPercent(topFiveOver45, activeOver45Total)} من إجمالي +45.</p>
								</article>
								<article {...widgetDragProps("owner-balances")} className={`rep-chart-card panel segment-panel dashboard-widget-draggable${visibleWidgets["owner-balances"] ? "" : " dashboard-widget-hidden"}`}>
								<div className="panel-heading">
									<div>
										<p className="panel-kicker">حسب المندوب</p>
										<h3>الرصيد حسب المندوب</h3>
									</div>
									<span className="panel-period">{activeSegments.length} مندوب</span>
								</div>
								<div className="segment-chart">
									{activeSegments.slice(0, 8).map(([representative, amount], index) => (
										<div className="segment-row" key={representative}>
											<div className="segment-label" title={representative}>
												<span className={`segment-dot segment-color-${index % 5}`} />
												<span>{representative}</span>
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
												<span>بقية المندوبين ({otherSegments.length})</span>
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
									<span>مندوبون غير محددين</span>
									<strong>{activeRepresentativeRows.filter((customer) => customer.segment === "غير محدد").length} حساب</strong>
								</div>
								</article>
							</div>
						</div>
						<article {...widgetDragProps("heatmap")} className={`rep-chart-card rep-heatmap-card dashboard-widget-draggable${visibleWidgets.heatmap ? "" : " dashboard-widget-hidden"}`}>
							<div className="rep-chart-heading">
								<div>
									<h3>خريطة أعمار العملاء · {activeRepresentativeName}</h3>
									<p>اضغط على عناوين الأعمدة للفرز · الأخضر حتى 45 يوماً والأحمر فوق 45</p>
								</div>
							</div>
							<div className="rep-heatmap-scroll">
								<table className="rep-heatmap">
									<thead>
										<tr>
											<th scope="col" aria-sort={heatmapSort.column === "customer" ? (heatmapSort.direction === "desc" ? "descending" : "ascending") : "none"}>
												<button className="sort-button" type="button" onClick={() => toggleHeatmapSort("customer")}>
													العميل / الحساب{heatmapSortIndicator("customer")}
												</button>
											</th>
											<th scope="col" aria-sort={heatmapSort.column === "segment" ? (heatmapSort.direction === "desc" ? "descending" : "ascending") : "none"}>
												<button className="sort-button" type="button" onClick={() => toggleHeatmapSort("segment")}>
													السيجمينت{heatmapSortIndicator("segment")}
												</button>
											</th>
											{dashboardResult?.bucketNames.map((bucket, index) => (
												<th
													scope="col"
													key={`${index}-${bucket}`}
													aria-sort={heatmapSort.column === "bucket" && heatmapSort.bucketIndex === index ? (heatmapSort.direction === "desc" ? "descending" : "ascending") : "none"}
												>
													<button className="sort-button" type="button" onClick={() => toggleHeatmapSort("bucket", index)}>
														{bucket}{heatmapSortIndicator("bucket", index)}
													</button>
												</th>
											))}
											<th scope="col" aria-sort={heatmapSort.column === "total" ? (heatmapSort.direction === "desc" ? "descending" : "ascending") : "none"}>
												<button className="sort-button" type="button" onClick={() => toggleHeatmapSort("total")}>
													الإجمالي{heatmapSortIndicator("total")}
												</button>
											</th>
										</tr>
									</thead>
									<tbody>
										{activeHeatmapRows.map((customer) => (
											<tr key={customer.account}>
												<th scope="row">
													<span className="rep-heatmap-customer-name">{customer.name || customer.account}</span>
													<small>{customer.account}</small>
												</th>
												<td className="rep-heatmap-segment">{customer.segment}</td>
												{dashboardResult?.bucketNames.map((bucket, index) => {
													const amount = customer.buckets[index] ?? 0;
													const intensity = Math.min(0.82, Math.max(0.08, Math.abs(amount) / maxCustomerBucket * 0.82));
													const riskCell = amount < 0 || index >= OVER_45_BUCKET_START;
													const warningCell = index >= OVER_30_BUCKET_START && index < OVER_45_BUCKET_START;
													return (
														<td
															key={`${customer.account}-${index}-${bucket}`}
															className={riskCell ? "heat-risk" : warningCell ? "heat-warning" : ""}
															style={{ backgroundColor: riskCell
																? `rgba(169, 79, 79, ${intensity})`
																: warningCell
																	? `rgba(230, 200, 79, ${intensity})`
																	: `rgba(61, 146, 114, ${intensity})` }}
															title={`${customer.name || customer.account} · ${bucket}: ${formatAmount(amount)}`}
														>
															{formatAmount(amount)}
														</td>
													);
												})}
												<td className="heat-total">{formatAmount(customer.total)}</td>
											</tr>
										))}
										{activeHeatmapRows.length === 0 && (
											<tr><td className="rep-heatmap-empty" colSpan={(dashboardResult?.bucketNames.length ?? 0) + 3}>لا توجد حسابات لهذا الاختيار.</td></tr>
										)}
									</tbody>
								</table>
							</div>
							<p className="rep-chart-footnote">{activeHeatmapRows.length} حساب · اضغط على مندوب آخر لعرض عملائه.</p>
						</article>
					</div>
				) : null}
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
												<strong>{archive.reportDate
													? `تاريخ الرصيد · ${formatReportDate(archive.reportDate)}`
													: formatReportDate(archive.reportDate)}</strong>
												<span>{archive.rowCount} حساب · رصيد {formatAmount(archive.total)}</span>
											</div>
											<button type="button" onClick={() => void removeArchive(archive)}>حذف النسخة</button>
										</div>
									))}
								</div>
							</>
						)}
						{comparison && (
							<div className="comparison-results">
								<p>
									المقارنة من {formatReportDate(comparison.first.reportDate)}
									{" إلى "}
									{formatReportDate(comparison.second.reportDate)}
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

					<div className="result-note">
						<span aria-hidden="true">i</span>
						نعرض فقط الأشخاص والسيجمينتات المحددة؛ أزل علامة الصح لإخفاء حساباتهم من الملخص والرسوم وملف Excel.
						{result.segmentMatches < result.rows.length && (
							<span> لم نجد سيجمينتاً لـ {result.rows.length - result.segmentMatches} حساب؛ تظهر كـ «غير محدد» ضمن بقية السيجمينتات.</span>
						)}
					</div>
				</section>
			)}

			</div>
			<section
				id="daily-sales-tab-panel"
				className="daily-sales-panel"
				role="tabpanel"
				aria-labelledby="daily-sales-tab"
				hidden={activeTab !== "daily-sales"}
			>
				<div className="daily-sales-heading">
					<div>
						<p className="panel-kicker">متابعة مباشرة من ملف Excel المشترك</p>
						<h1>متابعة المبيعات اليومية</h1>
					</div>
					<button
						className="button button-dark daily-sales-refresh"
						type="button"
						onClick={() => setDailySalesRefresh((value) => value + 1)}
					>
						<span aria-hidden="true">↻</span> تحديث البيانات
					</button>
				</div>
				<div className="daily-sales-toolbar">
					<div className="daily-sales-sheet-tabs" role="tablist" aria-label="أوراق متابعة المبيعات">
						{(["المبيعات اليومية", "متابعة الهدف"] as const).map((sheet) => (
							<button
								key={sheet}
								className={`daily-sales-sheet-tab${dailySalesSheet === sheet ? " active" : ""}`}
								type="button"
								role="tab"
								aria-selected={dailySalesSheet === sheet}
								onClick={() => setDailySalesSheet(sheet)}
							>
								{sheet}
							</button>
						))}
					</div>
					<a href={DAILY_SALES_WORKBOOK_URL} target="_blank" rel="noreferrer">
						فتح ملف Excel في OneDrive ↗
					</a>
				</div>
				<div className="daily-sales-workbook" key={`${dailySalesSheet}-${dailySalesRefresh}`}>
					<iframe
						title={`ورقة ${dailySalesSheet}`}
						src={dailySalesEmbedUrl(dailySalesSheet, dailySalesRefresh)}
						allow="clipboard-read; clipboard-write"
						referrerPolicy="no-referrer"
					/>
				</div>
				<p className="daily-sales-note" role="note">
					إذا لم تظهر الورقة داخل الصفحة، افتح الملف في OneDrive وتأكد من أن صلاحية الرابط تسمح بعرضه.
				</p>
			</section>
			<footer className="page-footer">
				<span>مِيزان<span className="footer-dot">.</span> تحليل الذمم ببساطة</span>
				<span>{activeTab === "dashboard" ? "معالجة محلية · بدون رفع بياناتك إلى خادم" : "بيانات مباشرة من ملف OneDrive المشترك"}</span>
			</footer>
		</main>
	);
}

export default App;
