CREATE TABLE IF NOT EXISTS archives (
	id TEXT PRIMARY KEY NOT NULL,
	created_at TEXT NOT NULL,
	total REAL NOT NULL,
	row_count INTEGER NOT NULL,
	segment_matches INTEGER NOT NULL,
	bucket_names_json TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS archive_rows (
	archive_id TEXT NOT NULL,
	account TEXT NOT NULL,
	name TEXT NOT NULL,
	customer_group TEXT NOT NULL,
	segment TEXT NOT NULL,
	buckets_json TEXT NOT NULL,
	total REAL NOT NULL,
	PRIMARY KEY (archive_id, account),
	FOREIGN KEY (archive_id) REFERENCES archives(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS archive_rows_segment_idx
	ON archive_rows (archive_id, segment);
