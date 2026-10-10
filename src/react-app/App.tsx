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
const AGING_CHART_COLORS = ["#3d9272", "#59a78c", "#78b89f", "#96c8b2", "#b1d4c3", "#d1bb6c", "#d69a55", "#cc7654", "#a94f4f"];
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

interface CollectionPriority {
	account: string;
	name: string;
	segment: string;
	total: number;
	over30: number;
	over45: number;
	previousOver45: number | null;
	changeOver45: number | null;
}

interface ArchiveSnapshotRows {
	id: string;
	createdAt: string;
	reportDate: string | null;
	rows: ProcessingResult["rows"];
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

const TREND_SNAPSHOT_LIMIT = 12;

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

function archiveOptionLabel(archive: ArchiveSummary): string {
	return formatReportDate(archive.reportDate);
}

function compareArchivesByBalanceDate(first: ArchiveSummary, second: ArchiveSummary): number {
	return (second.reportDate ?? second.createdAt.slice(0, 10)).localeCompare(first.reportDate ?? first.createdAt.slice(0, 10)) ||
		second.createdAt.localeCompare(first.createdAt);
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
	const [uploads, setUploads] = useState<Uploads>(EMPTY_UPLOADS);
	const [result, setResult] = useState<ProcessingResult | null>(null);
	const [lastUpdatedAt, setLastUpdatedAt] = useState<string | null>(null);
	const [currentArchiveId, setCurrentArchiveId] = useState<string | null>(null);
	const [previousSnapshot, setPreviousSnapshot] = useState<ProcessingResult["rows"] | null>(null);
	const [previousSnapshotAt, setPreviousSnapshotAt] = useState<string | null>(null);
	const [previousSnapshotReportDate, setPreviousSnapshotReportDate] = useState<string | null>(null);
	const [historySnapshots, setHistorySnapshots] = useState<ArchiveSnapshotRows[]>([]);
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
	const [activeOwnerId, setActiveOwnerId] = useState<string>("all");
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
					const history = await Promise.all(
						archiveList.slice(0, TREND_SNAPSHOT_LIMIT).map((archive) => loadArchive(archive.id)),
					);
					if (!isCurrent) return;
					const latest = history[0];
					const previous = history[1] ?? null;
					setResult(latest);
					setLastUpdatedAt(latest.createdAt);
					setCurrentArchiveId(latest.id);
					setPreviousSnapshot(previous?.rows ?? null);
					setPreviousSnapshotAt(previous?.createdAt ?? null);
					setPreviousSnapshotReportDate(previous?.reportDate ?? null);
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

	const dashboardResult = comparison?.second ?? result;
	const dashboardRows = useMemo(() => dashboardResult?.rows ?? [], [dashboardResult]);
	const dashboardPreviousRows = comparison?.first.rows ?? previousSnapshot;
	const dashboardPreviousDate = comparison
		? comparison.first.reportDate
		: previousSnapshotReportDate;
	const dashboardPreviousCreatedAt = comparison
		? comparison.first.createdAt
		: previousSnapshotAt;

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
	const activeRepresentativeUnder45 = activeRepresentativeRows.reduce(
		(sum, customer) => sum + sumBucketsThrough(customer, 6),
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
	const snapshotsForTrend = comparison
		? [
			{ id: comparison.first.id, createdAt: comparison.first.createdAt, reportDate: comparison.first.reportDate, rows: comparison.first.rows },
			{ id: comparison.second.id, createdAt: comparison.second.createdAt, reportDate: comparison.second.reportDate, rows: comparison.second.rows },
		]
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
	const previousOver45ByAccount = new Map(
		previousRowsForActive.map((customer) => [
			customer.account,
			sumBucketsFrom(customer, OVER_45_BUCKET_START),
		]),
	);
	const activeReducedOver45Count = activeRepresentativeRows.filter((customer) => {
		const previousOver45 = previousOver45ByAccount.get(customer.account);
		return previousOver45 !== undefined &&
			sumBucketsFrom(customer, OVER_45_BUCKET_START) < previousOver45;
	}).length;
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
	const maxRepresentativeBucket = Math.max(
		1,
		...representativeMetrics.flatMap((owner) =>
			(owner.rows[0]?.buckets ?? []).map((_, index) =>
				owner.rows.reduce((sum, customer) => sum + Math.max(0, customer.buckets[index] ?? 0), 0),
			),
		),
	);

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
	const priorityAccounts = useMemo<CollectionPriority[]>(() => {
		const previousByAccount = new Map(
			(previousSnapshot ?? []).map((customer) => [
				customer.account,
				sumBucketsFrom(customer, OVER_45_BUCKET_START),
			]),
		);
		return selectedRows
			.map((customer) => {
				const previousOver45 = previousByAccount.get(customer.account);
				const over45 = sumBucketsFrom(customer, OVER_45_BUCKET_START);
				return {
					account: customer.account,
					name: customer.name,
					segment: customer.segment,
					total: customer.total,
					over30: sumBucketsFrom(customer, OVER_30_BUCKET_START),
					over45,
					previousOver45: previousOver45 ?? null,
					changeOver45: previousOver45 === undefined ? null : over45 - previousOver45,
				};
			})
			.filter((customer) => customer.over30 > 0 || customer.over45 > 0)
			.sort((first, second) =>
				second.over45 - first.over45 ||
				second.over30 - first.over30 ||
				second.total - first.total,
			);
	}, [previousSnapshot, selectedRows]);
	const topPriorityAccounts = priorityAccounts.slice(0, 10);
	const topPriorityOver45 = topPriorityAccounts.reduce((sum, customer) => sum + customer.over45, 0);
	const newlyOver45Count = priorityAccounts.filter(
		(customer) => customer.previousOver45 !== null && customer.previousOver45 === 0 && customer.over45 > 0,
	).length;
	const increasedOver45Count = priorityAccounts.filter(
		(customer) => customer.changeOver45 !== null && customer.changeOver45 > 0,
	).length;

	function selectRepresentative(ownerId: string): void {
		setActiveOwnerId(ownerId);
		setSearch("");
		setSegmentFilter("all");
		setPage(1);
		if (ownerId === "all") {
			setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
			setShowOtherSegments(false);
		} else {
			setSelectedOwners(new Set([ownerId]));
			setShowOtherSegments(false);
		}
	}
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
	const formattedBalanceDate = result
		? formatReportDate(result.reportDate)
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
				if (result) {
					setPreviousSnapshot(result.rows);
					setPreviousSnapshotAt(lastUpdatedAt);
					setPreviousSnapshotReportDate(result.reportDate ?? null);
				}
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
					setArchives((current) => [archive, ...current].sort(compareArchivesByBalanceDate));
					setHistorySnapshots((current) => [
						{ id: archive.id, createdAt: archive.createdAt, reportDate: archive.reportDate, rows: processed.rows },
						...current.filter((snapshot) => snapshot.id !== archive.id),
					].sort((first, second) => compareArchivesByBalanceDate(first, second)).slice(0, TREND_SNAPSHOT_LIMIT));
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
			setArchives((current) => [archive, ...current].sort(compareArchivesByBalanceDate));
			setCurrentArchiveId(archive.id);
			setLastUpdatedAt(archive.createdAt);
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
			if (currentArchiveId === archive.id) {
				setCurrentArchiveId(remaining[0]?.id ?? null);
				if (remaining[0]) {
					const [latest, previous] = await Promise.all([
						loadArchive(remaining[0].id),
						remaining[1] ? loadArchive(remaining[1].id) : Promise.resolve(null),
					]);
					setResult(latest);
					setLastUpdatedAt(latest.createdAt);
					setPreviousSnapshot(previous?.rows ?? null);
					setPreviousSnapshotAt(previous?.createdAt ?? null);
					setPreviousSnapshotReportDate(previous?.reportDate ?? null);
				} else {
					setResult(null);
					setLastUpdatedAt(null);
					setPreviousSnapshot(null);
					setPreviousSnapshotAt(null);
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
					<span>تاريخ الرصيد</span>
					<strong>{isRestoring ? "جاري الاستعادة…" : formattedBalanceDate}</strong>
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

			<section className="sales-dashboard-hero" id="top" aria-labelledby="sales-dashboard-title">
				<div className="sales-dashboard-heading">
					<div>
						<p className="section-kicker">مؤشرات المحافظ والتحصيل</p>
						<h1 id="sales-dashboard-title">لوحة أداء المندوبين</h1>
						<p>اختر المندوب لمراجعة حجم المحفظة، أعمار الأرصدة، واتجاه التغير مقارنة بالنسخة السابقة.</p>
					</div>
				</div>
				<div className="dashboard-toolbar">
					<div className="sales-update-chip">
						<span className="privacy-dot" />
						<span>{dashboardResult
							? dashboardResult.reportDate
								? `تاريخ الرصيد · ${formatReportDate(dashboardResult.reportDate)}`
								: formatReportDate(dashboardResult.reportDate)
							: "بانتظار أول تحديث"}</span>
					</div>
					{archives.length > 0 && (
						<div className="comparison-controls dashboard-comparison-controls">
							<span className="comparison-strip-label">مقارنة:</span>
							<label>
								<span>من</span>
								<select
									aria-label="فترة الأساس للمقارنة"
									value={selectedArchiveIds[0]}
									onChange={(event) => {
										setSelectedArchiveIds((current) => [event.target.value, current[1]]);
										setComparison(null);
									}}
								>
									<option value="">اختر فترة</option>
									{archives.map((archive) => (
										<option key={archive.id} value={archive.id}>
											{archiveOptionLabel(archive)} · {formatAmount(archive.total)}
										</option>
									))}
								</select>
							</label>
							<label>
								<span>إلى</span>
								<select
									aria-label="فترة العرض للمقارنة"
									value={selectedArchiveIds[1]}
									onChange={(event) => {
										setSelectedArchiveIds(([first]) => [first, event.target.value]);
										setComparison(null);
									}}
								>
									<option value="">اختر فترة</option>
									{archives.map((archive) => (
										<option key={archive.id} value={archive.id}>
											{archiveOptionLabel(archive)} · {formatAmount(archive.total)}
										</option>
									))}
								</select>
							</label>
							<button
								className="button button-dark"
								type="button"
								disabled={isComparing || archives.length < 2 || !selectedArchiveIds[0] || !selectedArchiveIds[1] || selectedArchiveIds[0] === selectedArchiveIds[1]}
								onClick={() => void compareArchives()}
							>
								{isComparing ? "جارٍ…" : "تطبيق"}
							</button>
						</div>
					)}
				</div>
				<div className="representative-selector" role="group" aria-label="اختيار مندوب المبيعات">
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
							setSearch("");
							setSegmentFilter("all");
							setPage(1);
						}}
					>
						<span>باقي الفروع</span>
						<strong>{formatAmount(totalBalance(dashboardRows.filter((customer) => ownerForSegment(customer.segment) === null)))}</strong>
						<small>{dashboardRows.filter((customer) => ownerForSegment(customer.segment) === null).length} حساب</small>
					</button>
				</div>
				{result ? (
					<div
						key={`${activeOwnerId}-${showOtherSegments}-${[...selectedOwners].sort().join(",")}`}
						className="representative-view-transition"
					>
						<div className="representative-dashboard-heading">
							<div>
								<p className="panel-kicker">لوحة المندوب</p>
								<h2>{activeRepresentativeName}</h2>
							</div>
							<span className="representative-comparison-note">
								{dashboardPreviousCreatedAt
									? `مقارنة مع ${formatReportDate(dashboardPreviousDate)}`
									: "ارفع لقطة أخرى لإظهار اتجاه الأداء"}
							</span>
						</div>
						<div className="rep-kpi-grid">
							<article className="rep-kpi-card">
								<span>إجمالي الرصيد</span>
								<strong>{formatAmount(activeRepresentativeTotal)}</strong>
								<small className={`risk-change${activeChangeAmount === null || activeChangeAmount === 0 ? "" : activeChangeAmount > 0 ? " change-up" : " change-down"}`}>
									{activeChangeAmount === null
										? dashboardPreviousRows ? "لا توجد قيمة سابقة" : "لا توجد مقارنة بعد"
										: `${activeChangeAmount > 0 ? "↑ " : activeChangeAmount < 0 ? "↓ " : ""}${activeChangeAmount > 0 ? "+" : ""}${formatAmount(activeChangeAmount)} (${activeChangePercent === null ? "—" : `${activeChangePercent > 0 ? "+" : ""}${percentFormat.format(activeChangePercent)}٪`}) ${activeChangeAmount > 0 ? "· ارتفاع يحتاج متابعة" : activeChangeAmount < 0 ? "· انخفاض محتمل" : "· دون تغيير"}`}
								</small>
								<small className="rep-kpi-footnote">تغير الرصيد لا يثبت التحصيل وحده</small>
							</article>
							<article className="rep-kpi-card">
								<span>الرصيد حتى 21 يوماً</span>
								<strong>{formatAmount(activeRepresentativeUnder21)}</strong>
								<small>{formatPercent(activeRepresentativeUnder21, activeRepresentativeTotal)} من المحفظة · Current وحتى فترة أقل من 21</small>
							</article>
							<article className="rep-kpi-card">
								<span>الرصيد حتى 45 يوماً</span>
								<strong>{formatAmount(activeRepresentativeUnder45)}</strong>
								<small>{formatPercent(activeRepresentativeUnder45, activeRepresentativeTotal)} من المحفظة · الفترات حتى أقل من 45</small>
							</article>
							<article className="rep-kpi-card rep-kpi-risk">
								<span>الرصيد فوق 45 يوماً</span>
								<strong>{formatAmount(activeRepresentativeOver45)}</strong>
								<small>{formatPercent(activeRepresentativeOver45, activeRepresentativeTotal)} من المحفظة · {activeRepresentativeRows.filter((customer) => sumBucketsFrom(customer, OVER_45_BUCKET_START) > 0).length} حساب</small>
							</article>
							<article className="rep-kpi-card rep-kpi-risk">
								<span>العمولة المفقودة التقديرية · +45 · 1٪</span>
								<strong>{formatAmount(activeRepresentativeOver45 * COMMISSION_RATE)}</strong>
								<small>تقدير العمولة المرتبطة بالرصيد فوق 45 يوماً</small>
							</article>
							<article className="rep-kpi-card rep-kpi-warning">
								<span>العمولة المعرّضة للفقد · 30–45 · 1٪</span>
								<strong>{formatAmount(activeRepresentative30To45 * COMMISSION_RATE)}</strong>
								<small>الرصيد 30–45: {formatAmount(activeRepresentative30To45)}</small>
							</article>
							<article className="rep-kpi-card rep-kpi-warning">
								<span>حسابات دخلت +30 منذ فترة الأساس</span>
								<strong>{dashboardPreviousRows ? activeNewlyOver30Rows.length : "—"}</strong>
								<small>{dashboardPreviousRows ? `أرصدة +30 لهذه الحسابات: ${formatAmount(activeNewlyOver30Balance)}` : "اختر فترة مقارنة لعرض التغير"}</small>
							</article>
						</div>
						<div className="rep-chart-grid">
							<article className="rep-chart-card">
								<div className="rep-chart-heading">
									<div>
										<h3>اتجاه إجمالي الرصيد</h3>
										<p>{comparison ? "الفترة المحددة مقارنة بفترة الأساس" : `حسب آخر ${activeTrend.length} لقطات محفوظة`} · لا يمثل التحصيل وحده</p>
									</div>
									<strong>{formatAmount(activeRepresentativeTotal)}</strong>
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
											<span>{activeTrend[0] ? formatReportDate(activeTrend[0].reportDate) : ""}</span>
											<span>{activeTrend.length > 1 ? `${activeTrend.length} لقطات` : "لقطة واحدة"}</span>
											<span>{activeTrend.length > 1 ? formatReportDate(activeTrend[activeTrend.length - 1].reportDate) : ""}</span>
										</div>
									</>
								) : <p className="rep-chart-empty">لا توجد بيانات تاريخية لهذا المندوب.</p>}
							</article>
							<article className="rep-chart-card">
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
						<div className="rep-bottom-grid">
							<article className="rep-chart-card concentration-card">
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
							<article className="rep-chart-card">
								<div className="rep-chart-heading">
									<div>
										<h3>نبض المحفظة</h3>
										<p>إشارات متابعة تساعد على ترتيب يوم المندوب</p>
									</div>
								</div>
								<div className="rep-signal-list">
									<div><span>حسابات عليها رصيد +45</span><strong>{activeRepresentativeRows.filter((customer) => sumBucketsFrom(customer, OVER_45_BUCKET_START) > 0).length}</strong></div>
									<div className="signal-worsened"><span>دخلت +30 منذ فترة الأساس</span><strong>{dashboardPreviousRows ? activeNewlyOver30Rows.length : "—"}</strong></div>
									<div className="signal-improved"><span>انخفض رصيد +45 عن فترة الأساس</span><strong>{dashboardPreviousRows ? activeReducedOver45Count : "—"}</strong></div>
									<div className="signal-worsened"><span>عمولة معرضة للفقد · رصيد 30–45 · 1٪</span><strong>{formatAmount(activeRepresentative30To45 * COMMISSION_RATE)}</strong></div>
								</div>
							</article>
						</div>
						<article className="rep-chart-card rep-heatmap-card">
							<div className="rep-chart-heading">
								<div>
									<h3>خريطة أعمار الأرصدة حسب المندوب</h3>
								<p>قيمة كل فترة للمقارنة السريعة · الأخضر حتى 45 والأحمر +45</p>
								</div>
							</div>
							<div className="rep-heatmap-scroll">
								<table className="rep-heatmap">
									<thead>
										<tr>
											<th scope="col">المندوب</th>
											{dashboardResult?.bucketNames.map((bucket, index) => (
												<th scope="col" key={`${index}-${bucket}`}>{bucket}</th>
											))}
											<th scope="col">إجمالي +45</th>
										</tr>
									</thead>
									<tbody>
										{representativeMetrics.map((owner) => (
											<tr key={owner.id}>
												<th scope="row">{owner.name}</th>
												{dashboardResult?.bucketNames.map((bucket, index) => {
													const amount = owner.rows.reduce((sum, customer) => sum + (customer.buckets[index] ?? 0), 0);
													const intensity = Math.min(0.82, Math.max(0.08, Math.abs(amount) / maxRepresentativeBucket * 0.82));
													const riskCell = amount < 0 || index >= OVER_45_BUCKET_START;
													return (
														<td
															key={`${owner.id}-${index}-${bucket}`}
															className={riskCell ? "heat-risk" : ""}
															style={{ backgroundColor: riskCell ? `rgba(169, 79, 79, ${intensity})` : `rgba(61, 146, 114, ${intensity})` }}
															title={`${owner.name} · ${bucket}: ${formatAmount(amount)}`}
														>
															{formatAmount(amount)}
														</td>
													);
												})}
												<td className="heat-total">{formatAmount(owner.over45)}</td>
											</tr>
										))}
									</tbody>
								</table>
							</div>
						</article>
					</div>
				) : (
					<div className="rep-dashboard-empty">
						<div className="rep-empty-icon">↥</div>
						<h2>ارفع الملفات لبدء لوحة أداء المندوبين</h2>
						<p>اختر ملفي الفترات وملف السيجمينت من الشريط العلوي؛ ستظهر هنا المحافظ ومؤشرات أعمار الدين.</p>
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
										setActiveOwnerId("all");
										setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
										setShowOtherSegments(false);
										setPage(1);
									}}
								>
									المندوبون الستة فقط
								</button>
								<button
									type="button"
									onClick={() => {
										setActiveOwnerId("all");
										setSelectedOwners(new Set(OWNER_FILTERS.map((owner) => owner.id)));
										setShowOtherSegments(true);
										setPage(1);
									}}
								>
									مع باقي الفروع
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
												setActiveOwnerId("all");
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
										setActiveOwnerId("all");
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

					<section className="panel collection-priority-panel" aria-labelledby="collection-priority-title">
						<div className="panel-heading">
							<div>
								<p className="panel-kicker">ابدأ بالحسابات الأعلى تأثيراً</p>
								<h3 id="collection-priority-title">أولويات التحصيل حسب الرصيد المتأخر</h3>
							</div>
							<span className="panel-period">
								{previousSnapshotAt
									? `التغير مقابل ${formatReportDate(previousSnapshotReportDate)}`
									: "لا توجد نسخة سابقة لقياس التغير"}
							</span>
						</div>
						<div className="priority-insights" aria-live="polite">
							<div>
								<span>حصة أعلى 10 حسابات من +45</span>
								<strong>{formatPercent(topPriorityOver45, selectedOver45)}</strong>
								<small>{formatAmount(topPriorityOver45)} من {formatAmount(selectedOver45)}</small>
							</div>
							<div>
								<span>حسابات دخلت +45 منذ النسخة السابقة</span>
								<strong>{previousSnapshot ? newlyOver45Count : "—"}</strong>
								<small>{previousSnapshot ? "تستحق متابعة مبكرة" : "ارفع تحديثاً آخر لبدء المقارنة"}</small>
							</div>
							<div>
								<span>حسابات ارتفع لديها رصيد +45</span>
								<strong>{previousSnapshot ? increasedOver45Count : "—"}</strong>
								<small>{previousSnapshot ? "مقارنة بالقطة السابقة" : "لا يوجد خط أساس بعد"}</small>
							</div>
						</div>
						<div className="priority-table-scroll">
							<table className="priority-table">
								<thead>
									<tr>
										<th scope="col">#</th>
										<th scope="col">الحساب / العميل</th>
										<th scope="col">المندوب</th>
										<th scope="col">الرصيد الفعلي</th>
										<th scope="col">متأخر +30</th>
										<th scope="col">متأخر +45</th>
										<th scope="col">حصة +45 من المحفظة</th>
										<th scope="col">التغير +45</th>
										<th scope="col">إشارة متابعة</th>
									</tr>
								</thead>
								<tbody>
									{topPriorityAccounts.map((customer, index) => {
										const ownerTotal = commissionBySegment.find(([segment]) => segment === customer.segment)?.[1].over45 ?? 0;
										const signal = customer.previousOver45 === null
											? "أولوية حالية"
											: customer.previousOver45 === 0 && customer.over45 > 0
												? "دخل شريحة +45"
												: customer.changeOver45 !== null && customer.changeOver45 > 0
													? "ارتفع رصيد +45"
													: customer.changeOver45 !== null && customer.changeOver45 < 0
														? "انخفض رصيد +45"
														: "مستقر";
										return (
											<tr key={customer.account}>
												<td>{index + 1}</td>
												<td>
													<strong className="account-name">{customer.name || "—"}</strong>
													<span className="account-id">{customer.account}</span>
												</td>
												<td><span className="segment-tag">{customer.segment}</span></td>
												<td className="numeric-cell">{formatAmount(customer.total)}</td>
												<td className="numeric-cell">{formatAmount(customer.over30)}</td>
												<td className="numeric-cell priority-over45">{formatAmount(customer.over45)}</td>
												<td className="numeric-cell">{formatPercent(customer.over45, ownerTotal)}</td>
												<td className={`numeric-cell${customer.changeOver45 !== null && customer.changeOver45 > 0 ? " priority-increase" : customer.changeOver45 !== null && customer.changeOver45 < 0 ? " priority-decrease" : ""}`}>
													{customer.changeOver45 === null ? "—" : `${customer.changeOver45 > 0 ? "+" : ""}${formatAmount(customer.changeOver45)}`}
												</td>
												<td><span className={`priority-signal${signal.includes("دخل") || signal.includes("ارتفع") ? " warning" : signal.includes("انخفض") ? " easing" : ""}`}>{signal}</span></td>
											</tr>
										);
									})}
									{topPriorityAccounts.length === 0 && (
										<tr><td className="priority-empty" colSpan={9}>لا توجد أرصدة متأخرة ضمن الأشخاص والسيجمينتات المحددة.</td></tr>
									)}
								</tbody>
							</table>
						</div>
						<p className="priority-caveat">
							ترتيب الأولوية يعتمد على حجم الرصيد المتأخر، وليس على توقع التحصيل. انخفاض +45 لا يثبت أن المبلغ تم تحصيله؛ قد يتأثر بمبيعات أو مرتجعات أو تسويات بين النسختين.
						</p>
					</section>

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
