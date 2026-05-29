'use strict';

const { Pool } = require('pg');
const ExcelJS = require('exceljs');
const AWS = require('aws-sdk');

const DB_HOST = process.env.DB_HOST;
const DB_USER = process.env.userName;
const DB_PASS = process.env.DB_PASSWORD;
const DB_NAME = process.env.DB_NAME || 'AdvancedSearch';
const DB_PORT = parseInt(process.env.DB_PORT || '5432', 10);
const REPORT_TABLE = process.env.REPORT_TABLE;

const S3_BUCKET = process.env.S3_BUCKET;
const S3_PREFIX = process.env.S3_PREFIX || '';

const TYPE = 'bot-avg-response-time';

const basePgConfig = {
  host: DB_HOST,
  user: DB_USER,
  password: DB_PASS,
  port: DB_PORT,
  ssl: { rejectUnauthorized: false },
  max: 4,
  idleTimeoutMillis: 30000,
  connectionTimeoutMillis: 10000,
};

const poolAS = new Pool({ ...basePgConfig, database: DB_NAME });
const s3 = new AWS.S3();

function safeSheetName(name) {
  return String(name || 'Sheet').replace(/[\\/*?:[\]]/g, '_').slice(0, 31);
}

function autoFitColumns(worksheet) {
  worksheet.columns.forEach((column) => {
    let maxLength = 10;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const cellValue = cell.value == null ? '' : String(cell.value);
      maxLength = Math.max(maxLength, cellValue.length + 2);
    });
    column.width = Math.min(maxLength + 2, 50);
  });
}

function addSheetFromRows(workbook, sheetName, columns, rows) {
  const sheet = workbook.addWorksheet(safeSheetName(sheetName));
  sheet.columns = columns;
  sheet.getRow(1).font = { bold: true };
  for (const row of rows) sheet.addRow(row);
  autoFitColumns(sheet);
  return sheet;
}

function joinS3Key(prefix, key) {
  const p = String(prefix || '').trim();
  const k = String(key || '').trim();
  if (!p) return k;
  const p2 = p.endsWith('/') ? p.slice(0, -1) : p;
  const k2 = k.startsWith('/') ? k.slice(1) : k;
  return `${p2}/${k2}`;
}

function last12MonthsRangeDubai() {
  const now = new Date();

  const end = new Date(now);
  end.setHours(23, 59, 59, 999);

  const start = new Date(end);
  start.setFullYear(start.getFullYear() - 1);
  start.setHours(0, 0, 0, 0);

  return { startMs: start.getTime(), endMs: end.getTime() };
}

function formatDuration(seconds) {
  const s = Number(seconds ?? 0);
  if (!Number.isFinite(s) || s < 0) return '0 Sec';
  if (s < 60) return `${Math.round(s * 10) / 10} Sec`;

  const totalMinutes = Math.floor(s / 60);
  const remSeconds = Math.round(s % 60);

  if (totalMinutes < 60) {
    return remSeconds ? `${totalMinutes} Min ${remSeconds} Sec` : `${totalMinutes} Min`;
  }

  const totalHours = Math.floor(totalMinutes / 60);
  const remMinutes = totalMinutes % 60;

  if (totalHours < 24) {
    return remMinutes ? `${totalHours} Hour ${remMinutes} Min` : `${totalHours} Hour`;
  }

  const totalDays = Math.floor(totalHours / 24);
  const remHours = totalHours % 24;

  return remHours ? `${totalDays} Day ${remHours} Hour` : `${totalDays} Day ${0} Hour`;
}

async function getBotAvgResponseDailyRows(client, startMs, endMs) {
  const sql = `
    WITH days AS (
      SELECT generate_series(
        date_trunc('day', to_timestamp($1 / 1000.0))::date,
        date_trunc('day', to_timestamp($2 / 1000.0))::date,
        interval '1 day'
      )::date AS day
    ),
    per_ticket_day AS (
      SELECT
        date_trunc('day', to_timestamp("Created" / 1000.0))::date AS day,
        NULLIF(TRIM("a.ticket_id"), '') AS ticket_id,
        AVG(GREATEST(0, "Created" - "a.user_message_at"))::BIGINT AS avg_ms
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND "Created" IS NOT NULL
        AND "a.user_message_at" IS NOT NULL
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
      GROUP BY 1, 2
    ),
    per_day AS (
      SELECT
        day,
        COALESCE(AVG(avg_ms), 0)::BIGINT AS avg_ms
      FROM per_ticket_day
      GROUP BY day
    )
    SELECT
      d.day::text AS day,
      COALESCE(p.avg_ms, 0)::BIGINT AS avg_ms
    FROM days d
    LEFT JOIN per_day p
      ON p.day = d.day
    ORDER BY d.day ASC;
  `;

  const { rows } = await client.query(sql, [startMs, endMs]);
  return (rows || []).map((r) => ({
    day: String(r.day || ''),
    avg_ms: Number(r.avg_ms || 0),
  }));
}

async function buildWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();

  addSheetFromRows(
    workbook,
    'Chatbot Avg Response Time',
    [
      { header: 'Date', key: 'day' },
      { header: 'Average Bot Response Time', key: 'avgBotResponseTime' },
      { header: 'Average Seconds', key: 'avgSeconds' },
    ],
    rows.map((r) => {
      const avgSeconds = r.avg_ms ? Math.round((r.avg_ms / 1000) * 10) / 10 : 0;
      return {
        day: r.day,
        avgBotResponseTime: formatDuration(avgSeconds),
        avgSeconds,
      };
    })
  );

  return workbook;
}

async function uploadWorkbookToS3(workbook, key) {
  const buffer = await workbook.xlsx.writeBuffer();
  const body = Buffer.from(buffer);

  await s3
    .putObject({
      Bucket: S3_BUCKET,
      Key: key,
      Body: body,
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      CacheControl: 'no-store',
    })
    .promise();

  return { key, bytes: body.length };
}

exports.handler = async () => {
  if (!DB_HOST || !DB_USER || !DB_PASS || !REPORT_TABLE) {
    throw new Error('Missing required env vars: DB_HOST, userName, DB_PASSWORD, REPORT_TABLE');
  }
  if (!S3_BUCKET) {
    throw new Error('Missing required env var: S3_BUCKET');
  }

  const { startMs, endMs } = last12MonthsRangeDubai();

  const clientAS = await poolAS.connect();
  try {
    const rows = await getBotAvgResponseDailyRows(clientAS, startMs, endMs);
    const workbook = await buildWorkbook(rows);

    const filename = `${TYPE}.xlsx`;
    const s3Key = joinS3Key(S3_PREFIX, filename);

    const uploaded = await uploadWorkbookToS3(workbook, s3Key);

    return {
      ok: true,
      type: TYPE,
      startMs,
      endMs,
      bucket: S3_BUCKET,
      key: uploaded.key,
      bytes: uploaded.bytes,
    };
  } finally {
    clientAS.release();
  }
};