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

const TYPE = 'sla-assignment-time';

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

function formatDateTimeDubai(ms) {
  if (ms === undefined || ms === null || ms === '') return '';
  const n = Number(ms);
  if (!Number.isFinite(n)) return '';
  const d = new Date(n);
  if (Number.isNaN(d.getTime())) return '';

  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Dubai',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  }).formatToParts(d);

  const get = (type) => parts.find((p) => p.type === type)?.value || '';

  const yyyy = get('year');
  const mm = get('month');
  const dd = get('day');
  const hh = get('hour');
  const mi = get('minute');
  const ss = get('second');

  return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss} DXB`;
}

function formatDurationSeconds(seconds) {
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

async function getSlaDailyRows(client, startMs, endMs) {
  const sql = `
    WITH days AS (
      SELECT generate_series(
        date_trunc('day', to_timestamp($1 / 1000.0))::date,
        date_trunc('day', to_timestamp($2 / 1000.0))::date,
        interval '1 day'
      )::date AS day
    ),
    per_ticket AS (
      SELECT
        "a.ticket_id" AS ticket_id,
        date_trunc('day', to_timestamp(MAX("a.escalated_at") / 1000.0))::date AS day,
        MAX("a.escalated_at") AS escalated_at,
        MAX("a.assigned_to_agent_at") AS assigned_at
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
        AND "a.escalated_at" IS NOT NULL
        AND "a.assigned_to_agent_at" IS NOT NULL
      GROUP BY "a.ticket_id"
    ),
    filtered AS (
      SELECT
        day,
        ticket_id,
        escalated_at,
        assigned_at,
        GREATEST(0, assigned_at - escalated_at)::BIGINT AS duration_ms
      FROM per_ticket
      WHERE assigned_at >= escalated_at
    ),
    joined AS (
      SELECT
        d.day,
        f.ticket_id,
        f.escalated_at,
        f.assigned_at,
        f.duration_ms
      FROM days d
      LEFT JOIN filtered f
        ON f.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id,
      j.escalated_at,
      j.assigned_at,
      COALESCE(j.duration_ms, 0)::BIGINT AS duration_ms
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;

  const { rows } = await client.query(sql, [startMs, endMs]);
  return (rows || []).map((r) => ({
    day: String(r.day || ''),
    ticket_id: String(r.ticket_id || 'None'),
    escalated_at: r.escalated_at == null ? null : Number(r.escalated_at),
    assigned_at: r.assigned_at == null ? null : Number(r.assigned_at),
    duration_ms: Number(r.duration_ms || 0),
  }));
}

async function buildWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();

  addSheetFromRows(
    workbook,
    'SLA Assignment Time',
    [
      { header: 'Date', key: 'day' },
      { header: 'Ticket ID', key: 'ticketId' },
      { header: 'Escalation Time', key: 'escalationTime' },
      { header: 'Agent Assignment Time', key: 'assignmentTime' },
      { header: 'SLA Assignment Time', key: 'sla' },
    ],
    rows.map((r) => {
      const seconds = r.duration_ms ? Math.round((r.duration_ms / 1000) * 10) / 10 : 0;
      return {
        day: r.day,
        ticketId: r.ticket_id,
        escalationTime: r.escalated_at ? formatDateTimeDubai(r.escalated_at) : '',
        assignmentTime: r.assigned_at ? formatDateTimeDubai(r.assigned_at) : '',
        sla: r.ticket_id === 'None' ? '0 Sec' : formatDurationSeconds(seconds),
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
    const rows = await getSlaDailyRows(clientAS, startMs, endMs);
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