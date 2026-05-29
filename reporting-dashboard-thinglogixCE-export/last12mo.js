'use strict';

const AWS = require('aws-sdk');

const s3 = new AWS.S3();

const S3_BUCKET = process.env.S3_BUCKET;
const S3_PREFIX = (process.env.S3_PREFIX || '').trim();

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type,Authorization',
  'Access-Control-Allow-Methods': 'GET,OPTIONS',
};

function json(statusCode, bodyObj) {
  return {
    statusCode,
    headers: {
      ...corsHeaders,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(bodyObj),
  };
}

function okBinary(buffer, filename, contentType) {
  return {
    statusCode: 200,
    isBase64Encoded: true,
    headers: {
      ...corsHeaders,
      'Content-Type': contentType,
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Cache-Control': 'no-store',
      'Content-Length': String(buffer.length),
    },
    body: buffer.toString('base64'),
  };
}

function normalizePrefix(p) {
  if (!p) return '';
  return p.endsWith('/') ? p : `${p}/`;
}

const ALLOWED_TYPES = new Set([
  'assignments',
  'closed-tickets-by-chatbot',
  'handover-to-live-agent',
  'closed-assignment-rate',
  'bot-aht',
  'agent-aht',
  'bot-avg-response-time',
  'total-aht',
  'agent-first-response-time',
  'top-inquiry-types',
  'sla-assignment-time',
]);

async function getFileFromS3(type) {
  if (!S3_BUCKET) throw new Error('Missing required env var: S3_BUCKET');

  const t = String(type || '').trim();
  if (!ALLOWED_TYPES.has(t)) {
    return json(400, { error: `Unsupported type: ${t}` });
  }

  const key = `${normalizePrefix(S3_PREFIX)}${t}.xlsx`;

  try {
    const obj = await s3
      .getObject({
        Bucket: S3_BUCKET,
        Key: key,
      })
      .promise();

    const body = obj && obj.Body ? obj.Body : null;
    if (!body) return json(404, { error: 'File not found' });

    const buffer = Buffer.isBuffer(body) ? body : Buffer.from(body);

    const contentType =
      (obj.ContentType && String(obj.ContentType)) ||
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    const filename = `${t}.xlsx`;

    return okBinary(buffer, filename, contentType);
  } catch (err) {
    if (err && (err.code === 'NoSuchKey' || err.code === 'NotFound')) {
      return json(404, { error: 'File not found' });
    }
    console.error('S3 getObject error:', err);
    return json(500, { error: 'Failed to fetch file from S3' });
  }
}

module.exports = { getFileFromS3 };