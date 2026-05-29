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
const S3_PREFIX = (process.env.S3_PREFIX || 'last-12-months').replace(/^\/+|\/+$/g, '');

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
const s3 = new AWS.S3({ signatureVersion: 'v4' });

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

function safeSheetName(name) {
  return String(name || 'Sheet').replace(/[\\/*?:[\]]/g, '_').slice(0, 31);
}

function addSheetFromRows(workbook, sheetName, columns, rows) {
  const sheet = workbook.addWorksheet(safeSheetName(sheetName));
  sheet.columns = columns;
  sheet.getRow(1).font = { bold: true };

  for (const row of rows) sheet.addRow(row);

  autoFitColumns(sheet);
  return sheet;
}

async function getAssignmentsDailyRows(client, startMs, endMs) {
  const sql = `
    WITH bounds AS (
      SELECT
        date_trunc('day', timezone('Asia/Dubai', to_timestamp(($1::double precision) / 1000.0)))::date AS start_day,
        date_trunc('day', timezone('Asia/Dubai', to_timestamp(($2::double precision) / 1000.0)))::date AS end_day
    ),
    days AS (
      SELECT d::date AS day
      FROM bounds b
      CROSS JOIN generate_series(b.start_day, b.end_day, interval '1 day') AS d
    ),
    tickets AS (
      SELECT
        date_trunc(
          'day',
          timezone('Asia/Dubai', to_timestamp(("Created"::double precision) / 1000.0))
        )::date AS day,
        NULLIF(TRIM("a.ticket_id"), '') AS ticket_id
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
      GROUP BY 1, 2
    ),
    joined AS (
      SELECT
        to_char(days.day, 'YYYY-MM-DD') AS day_str,
        tickets.ticket_id AS ticket_id
      FROM days
      LEFT JOIN tickets
        ON tickets.day = days.day
    )
    SELECT
      day_str AS day,
      COALESCE(ticket_id, 'None') AS ticket_id
    FROM joined
    ORDER BY day_str ASC, (ticket_id IS NULL) ASC, ticket_id ASC;
  `;

  const { rows } = await client.query(sql, [startMs, endMs]);
  return (rows || []).map((r) => ({
    day: String(r.day || ''),
    ticket_id: String(r.ticket_id || ''),
  }));
}

async function buildAssignmentsWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();

  addSheetFromRows(
    workbook,
    'Total Assignments',
    [
      { header: 'Date', key: 'day' },
      { header: 'Ticket ID', key: 'ticketId' },
    ],
    rows.map((r) => ({
      day: r.day,
      ticketId: r.ticket_id,
    }))
  );

  return workbook;
}

function last12MonthsRangeMs(nowMs) {
  const dayMs = 24 * 60 * 60 * 1000;
  const endMs = Number(nowMs);
  const startMs = endMs - (365 * dayMs);
  return { startMs, endMs };
}

async function uploadWithAtomicSwap({ buffer, key }) {
  if (!S3_BUCKET) throw new Error('Missing required env var: S3_BUCKET');

  const tmpKey = `${S3_PREFIX}/_tmp/${key}`;
  const finalKey = `${S3_PREFIX}/${key}`;

  await s3.putObject({
    Bucket: S3_BUCKET,
    Key: tmpKey,
    Body: buffer,
    ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    CacheControl: 'no-store',
  }).promise();

  await s3.copyObject({
    Bucket: S3_BUCKET,
    CopySource: `${S3_BUCKET}/${tmpKey}`,
    Key: finalKey,
    ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    CacheControl: 'no-store',
    MetadataDirective: 'REPLACE',
  }).promise();

  await s3.deleteObject({
    Bucket: S3_BUCKET,
    Key: tmpKey,
  }).promise();

  return { bucket: S3_BUCKET, key: finalKey };
}

exports.handler = async () => {
  try {
    if (!DB_HOST || !DB_USER || !DB_PASS || !REPORT_TABLE) {
      throw new Error('Missing required env vars: DB_HOST, userName, DB_PASSWORD, REPORT_TABLE');
    }
    if (!S3_BUCKET) {
      throw new Error('Missing required env var: S3_BUCKET');
    }

    const nowMs = Date.now();
    const { startMs, endMs } = last12MonthsRangeMs(nowMs);

    const clientAS = await poolAS.connect();
    try {
      const rows = await getAssignmentsDailyRows(clientAS, startMs, endMs);
      const workbook = await buildAssignmentsWorkbook(rows);
      const xlsxBuffer = Buffer.from(await workbook.xlsx.writeBuffer());

      const uploaded = await uploadWithAtomicSwap({
        buffer: xlsxBuffer,
        key: 'assignments.xlsx',
      });

      return {
        ok: true,
        type: 'assignments',
        startMs,
        endMs,
        rows: rows.length,
        s3: uploaded,
      };
    } finally {
      clientAS.release();
    }
  } catch (err) {
    console.error('assignments-last-12-months error:', err);
    return {
      ok: false,
      error: err?.message || 'Internal server error',
    };
  }
};