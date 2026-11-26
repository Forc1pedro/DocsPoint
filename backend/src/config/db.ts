import { Pool } from 'pg';
import dotenv from 'dotenv';

dotenv.config();

const connectionString = process.env.DATABASE_URL;

export const pool = new Pool(
  connectionString
    ? { connectionString }
    : {
        host: process.env.POSTGRES_HOST || 'localhost',
        port: parseInt(process.env.POSTGRES_PORT || '5432', 10),
        user: process.env.POSTGRES_USER || 'sharepoint',
        password: process.env.POSTGRES_PASSWORD || 'sp_pass123',
        database: process.env.POSTGRES_DB || 'sharepoint_db',
      }
);

pool.on('connect', (client) => {
  client.query(
    'SET search_path TO content_storage, collaboration_lists, publishing_pages, governance_security, public;'
  ).catch((err) => {
    console.error('Failed to set search_path on PostgreSQL client', err);
  });
});

pool.on('error', (err) => {
  console.error('Unexpected error on idle PostgreSQL client', err);
});
