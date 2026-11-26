-- ==============================================================================
-- Migration: 002_separate_schemas.sql
-- Separate tables into domain schemas:
--  - content_storage
--  - collaboration_lists
--  - publishing_pages
--  - governance_security
-- ==============================================================================

CREATE SCHEMA IF NOT EXISTS content_storage;
CREATE SCHEMA IF NOT EXISTS collaboration_lists;
CREATE SCHEMA IF NOT EXISTS publishing_pages;
CREATE SCHEMA IF NOT EXISTS governance_security;

-- Move content & storage tables
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'sites') THEN
        ALTER TABLE public.sites SET SCHEMA content_storage;
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'document_libraries') THEN
        ALTER TABLE public.document_libraries SET SCHEMA content_storage;
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'documents') THEN
        ALTER TABLE public.documents SET SCHEMA content_storage;
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'document_versions') THEN
        ALTER TABLE public.document_versions SET SCHEMA content_storage;
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'document_shares') THEN
        ALTER TABLE public.document_shares SET SCHEMA content_storage;
    END IF;
END $$;

-- Move collaboration lists tables
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'lists') THEN
        ALTER TABLE public.lists SET SCHEMA collaboration_lists;
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'list_columns') THEN
        ALTER TABLE public.list_columns SET SCHEMA collaboration_lists;
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'list_items') THEN
        ALTER TABLE public.list_items SET SCHEMA collaboration_lists;
    END IF;
END $$;

-- Move publishing pages table
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'site_pages') THEN
        ALTER TABLE public.site_pages SET SCHEMA publishing_pages;
    END IF;
END $$;

-- Move governance & security tables
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'site_members') THEN
        ALTER TABLE public.site_members SET SCHEMA governance_security;
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'item_permissions') THEN
        ALTER TABLE public.item_permissions SET SCHEMA governance_security;
    END IF;
END $$;

-- Create Audit Logs table in governance_security
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

-- Set database default search_path for all future sessions
ALTER DATABASE sharepoint_db SET search_path TO content_storage, collaboration_lists, publishing_pages, governance_security, public;
