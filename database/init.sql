-- ==============================================================================
-- SharePoint Clone - Initial Database Schema
-- Database: PostgreSQL 16+
-- ==============================================================================

-- Enable UUID extension if required (built-in in PG16+, but safe fallback)
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- ==============================================================================
-- Domain Schemas Separation (TASK-03)
-- ==============================================================================
CREATE SCHEMA IF NOT EXISTS content_storage;
CREATE SCHEMA IF NOT EXISTS collaboration_lists;
CREATE SCHEMA IF NOT EXISTS publishing_pages;
CREATE SCHEMA IF NOT EXISTS governance_security;

ALTER DATABASE sharepoint_db SET search_path TO content_storage, collaboration_lists, publishing_pages, governance_security, public;
SET search_path TO content_storage, collaboration_lists, publishing_pages, governance_security, public;

-- ==============================================================================
-- 1. Sites
-- ==============================================================================
CREATE TABLE IF NOT EXISTS sites (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    title VARCHAR(255) NOT NULL,
    url_slug VARCHAR(100) UNIQUE NOT NULL,
    description TEXT,
    site_type VARCHAR(50) NOT NULL CHECK (site_type IN ('team', 'communication')),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Index for fast lookup by url_slug
CREATE INDEX IF NOT EXISTS idx_sites_url_slug ON sites(url_slug);

-- ==============================================================================
-- 2. Document Libraries (Bibliotecas de Documentos)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS document_libraries (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    site_id UUID NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
    name VARCHAR(255) NOT NULL,
    enable_versioning BOOLEAN DEFAULT TRUE,
    max_versions INT DEFAULT 50,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_site_library_name UNIQUE (site_id, name)
);

-- Index for libraries by site
CREATE INDEX IF NOT EXISTS idx_document_libraries_site_id ON document_libraries(site_id);

-- ==============================================================================
-- 3. Documents & Folders (Documentos e Pastas)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS documents (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    library_id UUID NOT NULL REFERENCES document_libraries(id) ON DELETE CASCADE,
    parent_folder_id UUID REFERENCES documents(id) ON DELETE CASCADE,
    name VARCHAR(255) NOT NULL,
    is_folder BOOLEAN DEFAULT FALSE,
    content_type VARCHAR(100),
    size_bytes BIGINT DEFAULT 0,
    current_version VARCHAR(20) DEFAULT '1.0',
    is_checked_out BOOLEAN DEFAULT FALSE,
    checked_out_by UUID,
    metadata JSONB DEFAULT '{}'::jsonb,
    deleted_at TIMESTAMP WITH TIME ZONE DEFAULT NULL,
    deleted_by UUID DEFAULT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Indexes for documents and folder navigation
CREATE INDEX IF NOT EXISTS idx_documents_library_id ON documents(library_id);
CREATE INDEX IF NOT EXISTS idx_documents_parent_folder ON documents(parent_folder_id);
CREATE INDEX IF NOT EXISTS idx_documents_name ON documents(name);
CREATE INDEX IF NOT EXISTS idx_documents_is_folder ON documents(is_folder);
CREATE INDEX IF NOT EXISTS idx_documents_metadata_gin ON documents USING gin (metadata);
CREATE INDEX IF NOT EXISTS idx_documents_deleted_at ON documents(deleted_at);

-- ==============================================================================
-- 4. Document Versions (Histórico de Versões)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS document_versions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    document_id UUID NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    version_number VARCHAR(20) NOT NULL,
    storage_path VARCHAR(500) NOT NULL,
    size_bytes BIGINT NOT NULL,
    comment TEXT,
    created_by UUID NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_document_version UNIQUE (document_id, version_number)
);

-- Index for versions by document
CREATE INDEX IF NOT EXISTS idx_document_versions_document_id ON document_versions(document_id);

-- ==============================================================================
-- 5. Document Shares (Compartilhamento Seguro)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS document_shares (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    document_id UUID NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    token VARCHAR(64) UNIQUE NOT NULL,
    permission VARCHAR(20) NOT NULL DEFAULT 'read' CHECK (permission IN ('read', 'edit')),
    created_by UUID,
    recipient_email VARCHAR(255),
    expires_at TIMESTAMP WITH TIME ZONE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_document_shares_token ON document_shares(token);
CREATE INDEX IF NOT EXISTS idx_document_shares_doc ON document_shares(document_id);

-- ==============================================================================
-- 5. Updated At Trigger Function
-- ==============================================================================
CREATE OR REPLACE FUNCTION update_timestamp_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$ LANGUAGE 'plpgsql';

DROP TRIGGER IF EXISTS trg_sites_updated_at ON sites;
CREATE TRIGGER trg_sites_updated_at
    BEFORE UPDATE ON sites
    FOR EACH ROW
    EXECUTE FUNCTION update_timestamp_column();

DROP TRIGGER IF EXISTS trg_document_libraries_updated_at ON document_libraries;
CREATE TRIGGER trg_document_libraries_updated_at
    BEFORE UPDATE ON document_libraries
    FOR EACH ROW
    EXECUTE FUNCTION update_timestamp_column();

DROP TRIGGER IF EXISTS trg_documents_updated_at ON documents;
CREATE TRIGGER trg_documents_updated_at
    BEFORE UPDATE ON documents
    FOR EACH ROW
    EXECUTE FUNCTION update_timestamp_column();


-- ==============================================================================
-- 6. Custom Lists (Épico 5)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS lists (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    site_id UUID NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
    title VARCHAR(255) NOT NULL,
    description TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS list_columns (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    list_id UUID NOT NULL REFERENCES lists(id) ON DELETE CASCADE,
    name VARCHAR(100) NOT NULL,
    display_name VARCHAR(255) NOT NULL,
    column_type VARCHAR(50) NOT NULL CHECK (column_type IN ('text', 'number', 'choice', 'date', 'boolean')),
    options JSONB DEFAULT '[]'::jsonb,
    is_required BOOLEAN DEFAULT FALSE,
    sort_order INT DEFAULT 0,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS list_items (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    list_id UUID NOT NULL REFERENCES lists(id) ON DELETE CASCADE,
    data JSONB NOT NULL DEFAULT '{}'::jsonb,
    created_by UUID,
    updated_by UUID,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_lists_site_id ON lists(site_id);
CREATE INDEX IF NOT EXISTS idx_list_columns_list_id ON list_columns(list_id);
CREATE INDEX IF NOT EXISTS idx_list_items_list_id ON list_items(list_id);
CREATE INDEX IF NOT EXISTS idx_list_items_data_gin ON list_items USING gin (data);

-- ==============================================================================
-- 7. Site Pages & Web Parts Canvas (Épico 6)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS site_pages (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    site_id UUID NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
    title VARCHAR(255) NOT NULL,
    slug VARCHAR(255) NOT NULL,
    description TEXT,
    canvas_content JSONB NOT NULL DEFAULT '[]'::jsonb,
    is_published BOOLEAN DEFAULT TRUE,
    created_by VARCHAR(255) DEFAULT 'Carlos Silva',
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT uq_site_page_slug UNIQUE (site_id, slug)
);

CREATE INDEX IF NOT EXISTS idx_site_pages_site_id ON site_pages(site_id);

-- ==============================================================================
-- 8. Site Members & Item Permissions (Épico 7)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS site_members (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    site_id UUID NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
    user_id UUID DEFAULT gen_random_uuid(),
    email VARCHAR(255) NOT NULL,
    name VARCHAR(255) NOT NULL,
    role VARCHAR(50) NOT NULL CHECK (role IN ('owner', 'member', 'visitor')),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS item_permissions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    item_id UUID NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    has_broken_inheritance BOOLEAN DEFAULT FALSE,
    user_email VARCHAR(255) NOT NULL,
    permission VARCHAR(50) NOT NULL CHECK (permission IN ('full_control', 'edit', 'read')),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_site_members_site_id ON site_members(site_id);
CREATE INDEX IF NOT EXISTS idx_item_permissions_item_id ON item_permissions(item_id);

-- ==============================================================================
-- 9. Audit Logs & Activity Feed (TASK-06)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS governance_security.audit_logs (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    site_id UUID,
    action VARCHAR(50) NOT NULL,
    target_type VARCHAR(50) NOT NULL,
    target_id UUID,
    target_name VARCHAR(255),
    user_name VARCHAR(255) DEFAULT 'Administrador do SharePoint',
    user_email VARCHAR(255) DEFAULT 'admin@sharepoint.local',
    details JSONB DEFAULT '{}'::jsonb,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_audit_logs_site_id ON governance_security.audit_logs(site_id);
CREATE INDEX IF NOT EXISTS idx_audit_logs_created_at ON governance_security.audit_logs(created_at DESC);
