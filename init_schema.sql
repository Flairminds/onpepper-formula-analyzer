-- init_schema.sql
-- Single-table schema for Excel Formula Analyzer

-- Table to store each workbook version with all related data in JSONB
CREATE TABLE IF NOT EXISTS workbook_records (
    record_id SERIAL PRIMARY KEY,
    file_name TEXT NOT NULL,
    version TEXT NOT NULL,
    uploaded_at TIMESTAMP WITH TIME ZONE DEFAULT now(),
    data JSONB NOT NULL   -- contains column mappings, sheets, formulas, dependencies, etc.
);

-- Indexes for fast lookup by file and version
CREATE INDEX IF NOT EXISTS idx_workbook_file_version ON workbook_records(file_name, version);
