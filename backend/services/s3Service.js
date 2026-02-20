import { S3Client, PutObjectCommand, GetObjectCommand } from '@aws-sdk/client-s3';

let s3ClientCache = null;

function getS3Client() {
    if (s3ClientCache) return s3ClientCache;
    const bucket = process.env.AWS_S3_BUCKET;
    if (!bucket) return null;
    const region = process.env.AWS_REGION || 'eu-north-1';
    s3ClientCache = new S3Client({
        region,
        credentials: {
            accessKeyId: process.env.AWS_ACCESS_KEY_ID || '',
            secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY || '',
        },
    });
    return s3ClientCache;
}

/**
 * Upload a buffer to S3 and return the object URL.
 * Reads AWS_S3_BUCKET at call time so env is correct after dotenv loads.
 * @param {Buffer} buffer - File buffer
 * @param {string} key - S3 object key (e.g. proposals/123/raw/file.xlsx)
 * @param {string} [contentType] - MIME type (default: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet for xlsx)
 * @returns {Promise<string|null>} Object URL or null if S3 not configured
 */
export async function uploadBufferToS3(buffer, key, contentType) {
    const bucket = process.env.AWS_S3_BUCKET;
    const client = getS3Client();
    if (!client || !bucket) {
        console.warn('[S3] AWS_S3_BUCKET not set; skipping upload');
        return null;
    }
    const region = process.env.AWS_REGION || 'eu-north-1';
    const mime = contentType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    await client.send(
        new PutObjectCommand({
            Bucket: bucket,
            Key: key,
            Body: buffer,
            ContentType: mime,
        })
    );
    const url = `https://${bucket}.s3.${region}.amazonaws.com/${key}`;
    return url;
}

/**
 * Fetch an object from S3 and return its body as a Buffer.
 * @param {string} key - S3 object key (e.g. proposals/123/raw/file.xlsx)
 * @returns {Promise<Buffer|null>} Buffer or null if not configured or object not found
 */
export async function getBufferFromS3(key) {
    const bucket = process.env.AWS_S3_BUCKET;
    const client = getS3Client();
    if (!client || !bucket) return null;
    try {
        const response = await client.send(
            new GetObjectCommand({ Bucket: bucket, Key: key })
        );
        const chunks = [];
        for await (const chunk of response.Body) {
            chunks.push(chunk);
        }
        return Buffer.concat(chunks);
    } catch (err) {
        if (err.name === 'NoSuchKey') return null;
        throw err;
    }
}

/**
 * Build the public object URL for a given key (same format as uploadBufferToS3).
 */
export function getS3UrlForKey(key) {
    const bucket = process.env.AWS_S3_BUCKET;
    if (!bucket) return null;
    const region = process.env.AWS_REGION || 'eu-north-1';
    return `https://${bucket}.s3.${region}.amazonaws.com/${key}`;
}
