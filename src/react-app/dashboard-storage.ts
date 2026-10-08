import type { ProcessingResult } from "./processor";

const DATABASE_NAME = "mizan-dashboard";
const STORE_NAME = "saved-dashboard";
const RECORD_KEY = "latest";

export interface SavedDashboard {
	result: ProcessingResult;
	updatedAt: string;
}

function openDatabase(): Promise<IDBDatabase> {
	if (!("indexedDB" in window)) {
		return Promise.reject(new Error("المتصفح لا يدعم التخزين المحلي المطلوب."));
	}

	return new Promise((resolve, reject) => {
		const request = window.indexedDB.open(DATABASE_NAME, 1);
		request.onupgradeneeded = () => {
			const database = request.result;
			if (!database.objectStoreNames.contains(STORE_NAME)) {
				database.createObjectStore(STORE_NAME);
			}
		};
		request.onsuccess = () => resolve(request.result);
		request.onerror = () => reject(request.error ?? new Error("تعذر فتح التخزين المحلي."));
		request.onblocked = () => reject(new Error("تعذر فتح التخزين المحلي بسبب تبويب آخر."));
	});
}

export async function loadSavedDashboard(): Promise<SavedDashboard | null> {
	const database = await openDatabase();
	try {
		return await new Promise((resolve, reject) => {
			const transaction = database.transaction(STORE_NAME, "readonly");
			const request = transaction.objectStore(STORE_NAME).get(RECORD_KEY);
			request.onsuccess = () => resolve((request.result as SavedDashboard | undefined) ?? null);
			request.onerror = () => reject(request.error ?? new Error("تعذر قراءة البيانات المحفوظة."));
			transaction.onabort = () => reject(transaction.error ?? new Error("تعذرت قراءة البيانات المحفوظة."));
		});
	} finally {
		database.close();
	}
}

export async function saveDashboard(record: SavedDashboard): Promise<void> {
	const database = await openDatabase();
	try {
		await new Promise<void>((resolve, reject) => {
			const transaction = database.transaction(STORE_NAME, "readwrite");
			transaction.objectStore(STORE_NAME).put(record, RECORD_KEY);
			transaction.oncomplete = () => resolve();
			transaction.onerror = () => reject(transaction.error ?? new Error("تعذر حفظ البيانات."));
			transaction.onabort = () => reject(transaction.error ?? new Error("تم إلغاء حفظ البيانات."));
		});
	} finally {
		database.close();
	}
}

export async function clearSavedDashboard(): Promise<void> {
	const database = await openDatabase();
	try {
		await new Promise<void>((resolve, reject) => {
			const transaction = database.transaction(STORE_NAME, "readwrite");
			transaction.objectStore(STORE_NAME).delete(RECORD_KEY);
			transaction.oncomplete = () => resolve();
			transaction.onerror = () => reject(transaction.error ?? new Error("تعذر مسح البيانات المحفوظة."));
			transaction.onabort = () => reject(transaction.error ?? new Error("تم إلغاء مسح البيانات."));
		});
	} finally {
		database.close();
	}
}
