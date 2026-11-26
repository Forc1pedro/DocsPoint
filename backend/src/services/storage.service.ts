import * as Minio from 'minio';
import dotenv from 'dotenv';
import { Readable } from 'stream';

dotenv.config();

const endpointUrl = process.env.S3_ENDPOINT || 'http://localhost:9000';
let endPoint = 'localhost';
let port = 9000;
let useSSL = false;

try {
  const parsed = new URL(endpointUrl);
  endPoint = parsed.hostname;
  port = parsed.port ? parseInt(parsed.port, 10) : parsed.protocol === 'https:' ? 443 : 80;
  useSSL = parsed.protocol === 'https:';
} catch {
  endPoint = process.env.S3_ENDPOINT?.replace(/^https?:\/\//, '') || 'localhost';
  port = parseInt(process.env.MINIO_API_PORT || '9000', 10);
  useSSL = process.env.S3_ENDPOINT?.startsWith('https') || false;
}

export class StorageService {
  private client: Minio.Client;
  public bucketName: string;

  constructor() {
    this.bucketName = process.env.S3_BUCKET || 'sp-documents';
    this.client = new Minio.Client({
      endPoint,
      port,
      useSSL,
      accessKey: process.env.S3_ACCESS_KEY || 'minioadmin',
      secretKey: process.env.S3_SECRET_KEY || 'minioadmin',
      region: process.env.S3_REGION || 'us-east-1',
    });
  }

  public async ensureBucket(): Promise<void> {
    try {
      const exists = await this.client.bucketExists(this.bucketName);
      if (!exists) {
        await this.client.makeBucket(this.bucketName, process.env.S3_REGION || 'us-east-1');
        console.log(`[StorageService] Bucket '${this.bucketName}' created successfully.`);
      }
    } catch (err) {
      console.error(`[StorageService] Error ensuring bucket '${this.bucketName}':`, err);
    }
  }

  public async uploadBuffer(
    objectName: string,
    buffer: Buffer,
    size: number,
    contentType: string,
    metaData: Record<string, string> = {}
  ): Promise<any> {
    await this.ensureBucket();
    return this.client.putObject(this.bucketName, objectName, buffer, size, {
      'Content-Type': contentType,
      ...metaData,
    });
  }

  public async uploadStream(
    objectName: string,
    stream: Readable,
    size: number,
    contentType: string,
    metaData: Record<string, string> = {}
  ): Promise<any> {
    await this.ensureBucket();
    return this.client.putObject(this.bucketName, objectName, stream, size, {
      'Content-Type': contentType,
      ...metaData,
    });
  }

  public async getPresignedUploadUrl(
    objectName: string,
    expirySeconds = 3600
  ): Promise<string> {
    await this.ensureBucket();
    return this.client.presignedPutObject(this.bucketName, objectName, expirySeconds);
  }

  public async getPresignedDownloadUrl(
    objectName: string,
    expirySeconds = 3600
  ): Promise<string> {
    await this.ensureBucket();
    return this.client.presignedGetObject(this.bucketName, objectName, expirySeconds);
  }

  public async getObjectStream(objectName: string): Promise<Readable> {
    return this.client.getObject(this.bucketName, objectName);
  }

  public async deleteObject(objectName: string): Promise<void> {
    return this.client.removeObject(this.bucketName, objectName);
  }
}

export const storageService = new StorageService();
