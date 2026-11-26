-- ==============================================================================
-- Migration: 003_tasks_07_to_16.sql
-- Database extensions for:
--  - TASK-09: Document Comments & Mentions
--  - TASK-10: Second-Stage Recycle Bin
--  - TASK-11: Custom Library Columns (Metadata)
--  - TASK-12: Document Approval Workflow
--  - TASK-15: Full-Text Content Search
-- ==============================================================================

SET search_path TO content_storage, collaboration_lists, publishing_pages, governance_security, public;

-- 1. Alter documents table for Approval, Second-stage recycle bin, and Content Search
ALTER TABLE content_storage.documents 
  ADD COLUMN IF NOT EXISTS content_text TEXT DEFAULT '',
  ADD COLUMN IF NOT EXISTS deleted_stage INT DEFAULT 1,
  ADD COLUMN IF NOT EXISTS approval_status VARCHAR(20) DEFAULT 'Approved',
  ADD COLUMN IF NOT EXISTS approval_comment TEXT,
  ADD COLUMN IF NOT EXISTS approved_by VARCHAR(255),
  ADD COLUMN IF NOT EXISTS approved_at TIMESTAMP WITH TIME ZONE;

-- Create GIN index for full-text search on documents content
CREATE INDEX IF NOT EXISTS idx_documents_content_text_fts 
  ON content_storage.documents USING gin (to_tsvector('portuguese', COALESCE(name, '') || ' ' || COALESCE(content_text, '')));

-- 2. Document Comments & Mentions table (TASK-09)
CREATE TABLE IF NOT EXISTS content_storage.document_comments (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    document_id UUID NOT NULL REFERENCES content_storage.documents(id) ON DELETE CASCADE,
    author_id UUID,
    author_name VARCHAR(255) NOT NULL,
    author_email VARCHAR(255) NOT NULL,
    content TEXT NOT NULL,
    mentions JSONB DEFAULT '[]'::jsonb,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_document_comments_doc_id ON content_storage.document_comments(document_id);
CREATE INDEX IF NOT EXISTS idx_document_comments_created_at ON content_storage.document_comments(created_at ASC);

-- 3. Custom Library Columns table (TASK-11)
CREATE TABLE IF NOT EXISTS content_storage.library_custom_columns (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    library_id UUID NOT NULL REFERENCES content_storage.document_libraries(id) ON DELETE CASCADE,
    name VARCHAR(100) NOT NULL,
    display_name VARCHAR(255) NOT NULL,
    column_type VARCHAR(50) NOT NULL CHECK (column_type IN ('text', 'choice', 'number', 'date', 'user')),
    options JSONB DEFAULT '[]'::jsonb,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_library_column_name UNIQUE (library_id, name)
);

CREATE INDEX IF NOT EXISTS idx_library_custom_columns_lib ON content_storage.library_custom_columns(library_id);
