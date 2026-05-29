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

const TYPE = 'closed-assignment-rate';

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

async function getClosedAssignmentRateDaily(client, startMs, endMs) {
  const sql = `
    WITH days AS (
      SELECT generate_series(
        date_trunc('day', to_timestamp($1 / 1000.0))::date,
        date_trunc('day', to_timestamp($2 / 1000.0))::date,
        interval '1 day'
      )::date AS day
    ),
    base AS (
      SELECT
        date_trunc('day', to_timestamp("Created" / 1000.0))::date AS day,
        NULLIF(TRIM("a.ticket_id"), '') AS ticket_id,
        COALESCE("a.close_status", '') AS close_status,
        "a.assigned_to_agent_at" AS assigned_to_agent_at,
        COALESCE("a.answer", '') AS answer,
        "a.escalated" AS escalated
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
    ),
    assignments_daily AS (
      SELECT day, COUNT(DISTINCT ticket_id)::BIGINT AS assignments
      FROM base
      GROUP BY day
    ),
    closed_by_chatbot_daily AS (
      WITH latest_per_ticket AS (
        SELECT DISTINCT ON (ticket_id)
          day,
          ticket_id,
          close_status,
          assigned_to_agent_at,
          answer
        FROM base
        ORDER BY ticket_id, day DESC
      )
      SELECT day, COUNT(DISTINCT ticket_id)::BIGINT AS closed_by_chatbot
      FROM latest_per_ticket
      WHERE close_status <> 'NOT_CLOSED'
        AND assigned_to_agent_at IS NULL
        AND answer NOT ILIKE '%now connecting you to a live agent. Please wait a moment while I transfer you%'
      GROUP BY day
    ),
    handover_daily AS (
      SELECT day, COUNT(DISTINCT ticket_id)::BIGINT AS handed_over
      FROM base
      WHERE escalated IS TRUE
      GROUP BY day
    )
    SELECT
      d.day::text AS day,
      COALESCE(a.assignments, 0) AS assignments,
      COALESCE(c.closed_by_chatbot, 0) AS closed_by_chatbot,
      COALESCE(h.handed_over, 0) AS handed_over,
      CASE
        WHEN COALESCE(a.assignments, 0) > 0
          THEN ROUND((COALESCE(c.closed_by_chatbot, 0)::numeric / a.assignments::numeric) * 100, 1)
        ELSE 0
      END AS closed_by_chatbot_pct
    FROM days d
    LEFT JOIN assignments_daily a ON a.day = d.day
    LEFT JOIN closed_by_chatbot_daily c ON c.day = d.day
    LEFT JOIN handover_daily h ON h.day = d.day
    ORDER BY d.day ASC;
  `;

  const { rows } = await client.query(sql, [startMs, endMs]);
  return (rows || []).map((r) => ({
    day: String(r.day || ''),
    assignments: Number(r.assignments || 0),
    closed_by_chatbot: Number(r.closed_by_chatbot || 0),
    handed_over: Number(r.handed_over || 0),
    closed_by_chatbot_pct: Number(r.closed_by_chatbot_pct || 0),
  }));
}

async function buildWorkbook(dailyRows) {
  const workbook = new ExcelJS.Workbook();

  addSheetFromRows(
    workbook,
    'Closed Assignment Rate',
    [
      { header: 'Date', key: 'day' },
      { header: 'Assignments', key: 'assignments' },
      { header: 'Closed by Chatbot', key: 'closed_by_chatbot' },
      { header: 'Handed Over', key: 'handed_over' },
      { header: 'Closed by Chatbot %', key: 'closed_by_chatbot_pct' },
    ],
    dailyRows.map((r) => ({
      day: r.day,
      assignments: r.assignments,
      closed_by_chatbot: r.closed_by_chatbot,
      handed_over: r.handed_over,
      closed_by_chatbot_pct: `${r.closed_by_chatbot_pct}%`,
    }))
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
    const dailyRows = await getClosedAssignmentRateDaily(clientAS, startMs, endMs);
    const workbook = await buildWorkbook(dailyRows);

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