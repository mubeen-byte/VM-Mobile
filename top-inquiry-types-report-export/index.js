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

const TYPE = 'top-inquiry-types';

const INQUIRY_CATEGORIES = [
  "One country roaming enquiry",
  "One country roaming issue",
  "One country calling issue",
  "One country calling enquiry",
  "Iphone 17 - pre-order issue",
  "Delayed delivery - Complaint",
  "Iphone 17 - Enquiry",
  "Minors SIM",
  "UAE PASS - eSIM Signup",
  "Calls Barred-provisioning issue",
  "6 and 12 months renewal refund",
  "iPhone - Refund",
  "iPhone Exchange Request",
  "Refund - Renewal Mobile",
  "Refund - Renewal MBB",
  "Groups - Complaints",
  "Groups - Benefits",
  "Groups - How it works",
  "DNCR",
  "Change email - Blocked CX",
  "Refund - Auto renewal complaint",
  "Entertainer Query-Complaint",
  "ID - Update Issue - UAE pass",
  "Duplicate-Incorrect charge",
  "Roaming Pass",
  "MBB-Cost",
  "Confirm payment query",
  "Find account details",
  "Refund - Within 3 days",
  "VIP-VVIP Plans",
  "Unblock lost device",
  "Block lost device",
  "React Rqst-Termin due ID updated",
  "SWAP Number",
  "Family Promo not added",
  "Adjust Free Renewals",
  "Number Activation Issue",
  "Notification Issue",
  "App Password Reset link Issue",
  "Not able to add card",
  "Login Issue",
  "MNP - Port out issues",
  "MNP - Port in Issues",
  "ID - Reconnection Issue",
  "ID - Update Issue",
  "Booster not added",
  "Plan Not Renewed",
  "Plan not added",
  "Hard Sim Capping",
  "Voice Spam",
  "MBB-Exchange",
  "MBB-Data Issue",
  "MBB-Refund-Return",
  "MBB-General Query",
  "Botim VOIP",
  "Fraud Blocks",
  "Tourist Plans Query",
  "Follow-up on tickets",
  "Data Rollover",
  "Double Data",
  "Promoter overpromise - Cashback",
  "Anghami",
  "SIM query",
  "5G-4G issue",
  "Family Plan Inquiry",
  "People of Determination",
  "Premium SMS",
  "TDRA Complaint",
  "Loyalty",
  "Vouchers",
  "Stores- Locations",
  "Credit Sharing",
  "Data Sharing",
  "Internet Calling",
  "Usage Details",
  "Manager Call Back",
  "Promo Code not added",
  "eSIM",
  "Add Sim Issue",
  "Apple Watch",
  "VAT Invoice",
  "Change Customers Details",
  "Change Of Ownership",
  "Prospective Customer",
  "Offers -Promotions",
  "Staff Complaint",
  "Cancellation",
  "Lost Sim",
  "ID Update -Expiry- Status query",
  "Calls Unable make -receive",
  "SIM Delivery",
  "Roaming",
  "Coverage -Network Issues -Data",
  "SMS Issue",
  "Plan Update",
  "MNP Query",
  "Booster-Pause query",
  "Wallet-Recharge",
  "App Errors Bug Crashes",
  "Re-activation - Activation query",
  "Account Email query",
  "Live Agent Handover",
  "Greeting",
  "Out of Scope"
];

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

async function getTopInquiryTypesDailyRows(client, startMs, endMs) {
  const sql = `
    WITH days AS (
      SELECT generate_series(
        date_trunc('day', to_timestamp($2 / 1000.0))::date,
        date_trunc('day', to_timestamp($3 / 1000.0))::date,
        interval '1 day'
      )::date AS day
    ),
    categories AS (
      SELECT unnest($1::text[]) AS inquiry_type
    ),
    base_rows AS (
      SELECT
        date_trunc('day', to_timestamp("Created" / 1000.0))::date AS day,
        NULLIF(TRIM("a.inquiry_type"), '') AS inquiry_type,
        NULLIF(TRIM("a.ticket_id"), '') AS ticket_id
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $2 AND $3
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
    ),
    normal_rows AS (
      SELECT
        day,
        inquiry_type,
        ticket_id
      FROM base_rows
      WHERE inquiry_type IS NOT NULL
        AND inquiry_type <> 'Live Agent Handover'
    ),
    live_agent_rows AS (
      SELECT
        day,
        'Live Agent Handover'::text AS inquiry_type,
        ticket_id
      FROM base_rows
      WHERE inquiry_type = 'Live Agent Handover'
    ),
    combined AS (
      SELECT * FROM normal_rows
      UNION ALL
      SELECT * FROM live_agent_rows
    ),
    counts AS (
      SELECT
        day,
        inquiry_type,
        COUNT(*)::BIGINT AS count
      FROM combined
      GROUP BY day, inquiry_type
    ),
    matched_rows AS (
      SELECT
        c.day,
        c.inquiry_type,
        c.ticket_id,
        cnt.count
      FROM combined c
      JOIN counts cnt
        ON cnt.day = c.day
       AND cnt.inquiry_type = c.inquiry_type
    ),
    zero_rows AS (
      SELECT
        d.day,
        cat.inquiry_type,
        'None'::text AS ticket_id,
        0::BIGINT AS count
      FROM days d
      CROSS JOIN categories cat
      LEFT JOIN counts cnt
        ON cnt.day = d.day
       AND cnt.inquiry_type = cat.inquiry_type
      WHERE COALESCE(cnt.count, 0) = 0
    ),
    final_rows AS (
      SELECT day, inquiry_type, ticket_id, count FROM matched_rows
      UNION ALL
      SELECT day, inquiry_type, ticket_id, count FROM zero_rows
    )
    SELECT
      fr.day::text AS day,
      fr.inquiry_type,
      fr.ticket_id,
      fr.count
    FROM final_rows fr
    ORDER BY
      fr.day ASC,
      fr.count DESC,
      fr.inquiry_type ASC,
      (fr.ticket_id = 'None') ASC,
      fr.ticket_id ASC;
  `;

  const { rows } = await client.query(sql, [INQUIRY_CATEGORIES, startMs, endMs]);
  return (rows || []).map((r) => ({
    day: String(r.day || ''),
    inquiry_type: String(r.inquiry_type || ''),
    ticket_id: String(r.ticket_id || 'None'),
    count: Number(r.count || 0),
  }));
}

async function buildWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();

  addSheetFromRows(
    workbook,
    'FAQs',
    [
      { header: 'Date', key: 'day' },
      { header: 'Question/Inquiry', key: 'inquiry' },
      { header: 'Ticket ID', key: 'ticketId' },
      { header: 'Count', key: 'count' },
    ],
    rows.map((r) => ({
      day: r.day,
      inquiry: r.inquiry_type,
      ticketId: r.ticket_id,
      count: r.count,
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
    const rows = await getTopInquiryTypesDailyRows(clientAS, startMs, endMs);
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