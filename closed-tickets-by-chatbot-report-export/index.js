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
const S3_PREFIX = (process.env.S3_PREFIX || '').replace(/^\/+|\/+$/g, '');

const TYPE = 'closed-tickets-by-chatbot';

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

function formatFileDate(ms) {
  const d = new Date(Number(ms));
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

async function getClosedTicketsDailyRows(client, startMs, endMs) {
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
    latest_per_ticket AS (
      SELECT DISTINCT ON ("a.ticket_id")
        "a.ticket_id" AS ticket_id,
        "a.close_status",
        "a.assigned_to_agent_at",
        "a.answer",
        "Created" AS created_ms
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
      ORDER BY "a.ticket_id", "Created" DESC
    ),
    closed_by_bot AS (
      SELECT
        date_trunc(
          'day',
          timezone('Asia/Dubai', to_timestamp((created_ms::double precision) / 1000.0))
        )::date AS day,
        NULLIF(TRIM(ticket_id), '') AS ticket_id
      FROM latest_per_ticket
      WHERE COALESCE("a.close_status", '') <> 'NOT_CLOSED'
        AND "a.assigned_to_agent_at" IS NULL
        AND COALESCE("a.answer",'') NOT ILIKE '%now connecting you to a live agent. Please wait a moment while I transfer you%'
    ),
    joined AS (
      SELECT
        to_char(days.day, 'YYYY-MM-DD') AS day_str,
        closed_by_bot.ticket_id AS ticket_id
      FROM days
      LEFT JOIN closed_by_bot
        ON closed_by_bot.day = days.day
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

async function buildWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();
  addSheetFromRows(
    workbook,
    'Closed by Chatbot',
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

async function uploadWorkbookToS3(workbook, key) {
  const buffer = await workbook.xlsx.writeBuffer();
  await s3
    .putObject({
      Bucket: S3_BUCKET,
      Key: key,
      Body: Buffer.from(buffer),
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      CacheControl: 'no-store',
    })
    .promise();
}

exports.handler = async () => {
  if (!DB_HOST || !DB_USER || !DB_PASS || !REPORT_TABLE) {
    throw new Error('Missing required env vars: DB_HOST, userName, DB_PASSWORD, REPORT_TABLE');
  }
  if (!S3_BUCKET) {
    throw new Error('Missing required env vars: S3_BUCKET');
  }

  const nowMs = Date.now();
  const dayMs = 24 * 60 * 60 * 1000;

  const endMs = nowMs;
  const startMs = endMs - (364 * dayMs);

  const clientAS = await poolAS.connect();
  try {
    const rows = await getClosedTicketsDailyRows(clientAS, startMs, endMs);
    const workbook = await buildWorkbook(rows);

    const keyBase = S3_PREFIX ? `${S3_PREFIX}/${TYPE}.xlsx` : `${TYPE}.xlsx`;
    await uploadWorkbookToS3(workbook, keyBase);

    return {
      ok: true,
      type: TYPE,
      s3_bucket: S3_BUCKET,
      s3_key: keyBase,
      start: formatFileDate(startMs),
      end: formatFileDate(endMs),
      rows: rows.length,
    };
  } finally {
    clientAS.release();
  }
};