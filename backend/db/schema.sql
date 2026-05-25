CREATE TABLE IF NOT EXISTS reference_sheets (
    kind TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    source_file TEXT,
    sheet_order INTEGER NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (kind, sheet_name)
);

CREATE TABLE IF NOT EXISTS reference_rows (
    id BIGSERIAL PRIMARY KEY,
    kind TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    row_index INTEGER NOT NULL,
    payload JSONB NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT fk_reference_sheet
        FOREIGN KEY (kind, sheet_name)
        REFERENCES reference_sheets (kind, sheet_name)
        ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_reference_rows_lookup
    ON reference_rows (kind, sheet_name, row_index);
