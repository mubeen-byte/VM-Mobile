'use strict';

const { Pool } = require('pg');
const ExcelJS = require('exceljs');

const DB_HOST = process.env.DB_HOST;
const DB_USER = process.env.userName;
const DB_PASS = process.env.DB_PASSWORD;
const DB_NAME = process.env.DB_NAME || 'AdvancedSearch';
const DB_PORT = parseInt(process.env.DB_PORT || '5432', 10);
const REPORT_TABLE = process.env.REPORT_TABLE;
const LIVE_AGENT_TABLE = process.env.LIVE_AGENT_TABLE || 'public.live_agent_data';

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
const poolPG = new Pool({ ...basePgConfig, database: 'postgres' });

const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization',
    'Access-Control-Allow-Methods': 'GET,OPTIONS',
};

function json(statusCode, bodyObj) {
    return {
        statusCode,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
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

function last24HoursRangeDubai() {
    const now = new Date();
    const end = new Date(now);
    const start = new Date(end.getTime() - 24 * 60 * 60 * 1000);
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

async function workbookBufferToResponse(workbook, filename) {
    const buffer = await workbook.xlsx.writeBuffer();
    return okBinary(
        Buffer.from(buffer),
        filename,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
}

async function getAssignmentsDailyRows(client, startMs, endMs) {
    const sql = `
    WITH days AS (
      SELECT generate_series(
        date_trunc('day', to_timestamp($1 / 1000.0))::date,
        date_trunc('day', to_timestamp($2 / 1000.0))::date,
        interval '1 day'
      )::date AS day
    ),
    tickets AS (
      SELECT
        date_trunc('day', to_timestamp("Created" / 1000.0))::date AS day,
        to_char((to_timestamp("Created" / 1000.0) AT TIME ZONE 'Asia/Dubai'), 'YYYY-MM-DD HH24:MI:SS') || ' DXB' AS date_time,
        NULLIF(TRIM("a.ticket_id"), '') AS ticket_id
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
      GROUP BY 1, 2, 3
    ),
    joined AS (
      SELECT d.day, t.ticket_id
      FROM days d
      LEFT JOIN tickets t ON t.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;
    const { rows } = await client.query(sql, [startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        ticket_id: String(r.ticket_id || 'None'),
    }));
}

async function getClosedByChatbotDailyRows(client, startMs, endMs) {
    const sql = `
    WITH days AS (
      SELECT generate_series(
        date_trunc('day', to_timestamp($1 / 1000.0))::date,
        date_trunc('day', to_timestamp($2 / 1000.0))::date,
        interval '1 day'
      )::date AS day
    ),
    latest_per_ticket AS (
      SELECT DISTINCT ON ("a.ticket_id")
        "a.ticket_id" AS ticket_id,
        "a.close_status",
        "a.assigned_to_agent_at",
        "a.answer",
        "Created"
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
      ORDER BY "a.ticket_id", "Created" DESC
    ),
    closed AS (
      SELECT
        date_trunc('day', to_timestamp(l."Created" / 1000.0))::date AS day,
        l.ticket_id
      FROM latest_per_ticket l
      WHERE COALESCE(l."a.close_status", '') <> 'NOT_CLOSED'
        AND l."a.assigned_to_agent_at" IS NULL
        AND COALESCE(l."a.answer",'') NOT ILIKE '%now connecting you to a live agent. Please wait a moment while I transfer you%'
    ),
    joined AS (
      SELECT d.day, c.ticket_id
      FROM days d
      LEFT JOIN closed c ON c.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;
    const { rows } = await client.query(sql, [startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        ticket_id: String(r.ticket_id || 'None'),
    }));
}

async function getHandoverDailyRows(client, startMs, endMs) {
    const sql = `
    WITH days AS (
      SELECT generate_series(
        date_trunc('day', to_timestamp($1 / 1000.0))::date,
        date_trunc('day', to_timestamp($2 / 1000.0))::date,
        interval '1 day'
      )::date AS day
    ),
    tickets AS (
      SELECT
        date_trunc('day', to_timestamp("Created" / 1000.0))::date AS day,
        to_char((to_timestamp("Created" / 1000.0) AT TIME ZONE 'Asia/Dubai'), 'YYYY-MM-DD HH24:MI:SS') || ' DXB' AS date_time,
        NULLIF(TRIM("a.ticket_id"), '') AS ticket_id
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
        AND COALESCE("a.escalated", FALSE) IS TRUE
      GROUP BY 1, 2, 3
    ),
    joined AS (
      SELECT d.day, t.ticket_id
      FROM days d
      LEFT JOIN tickets t ON t.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;
    const { rows } = await client.query(sql, [startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        ticket_id: String(r.ticket_id || 'None'),
    }));
}

async function getClosedAssignmentRateDailyRows(client, startMs, endMs) {
    const sql = `
    WITH days AS (
      SELECT generate_series(
        date_trunc('day', to_timestamp($1 / 1000.0))::date,
        date_trunc('day', to_timestamp($2 / 1000.0))::date,
        interval '1 day'
      )::date AS day
    ),
    assignments AS (
      SELECT
        date_trunc('day', to_timestamp("Created" / 1000.0))::date AS day,
        COUNT(DISTINCT NULLIF(TRIM("a.ticket_id"), ''))::BIGINT AS assignments
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
      GROUP BY 1
    ),
    closed_latest AS (
      SELECT DISTINCT ON ("a.ticket_id")
        "a.ticket_id" AS ticket_id,
        "a.close_status",
        "a.assigned_to_agent_at",
        "a.answer",
        "Created"
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
      ORDER BY "a.ticket_id", "Created" DESC
    ),
    closed_by_chatbot AS (
      SELECT
        date_trunc('day', to_timestamp(l."Created" / 1000.0))::date AS day,
        COUNT(*)::BIGINT AS closed_by_chatbot
      FROM closed_latest l
      WHERE COALESCE(l."a.close_status", '') <> 'NOT_CLOSED'
        AND l."a.assigned_to_agent_at" IS NULL
        AND COALESCE(l."a.answer",'') NOT ILIKE '%now connecting you to a live agent. Please wait a moment while I transfer you%'
      GROUP BY 1
    ),
    handed_over AS (
      SELECT
        date_trunc('day', to_timestamp("Created" / 1000.0))::date AS day,
        COUNT(DISTINCT NULLIF(TRIM("a.ticket_id"), ''))::BIGINT AS handed_over
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
        AND COALESCE("a.escalated", FALSE) IS TRUE
      GROUP BY 1
    )
    SELECT
      d.day::text AS day,
      COALESCE(a.assignments, 0)::BIGINT AS assignments,
      COALESCE(c.closed_by_chatbot, 0)::BIGINT AS closed_by_chatbot,
      COALESCE(h.handed_over, 0)::BIGINT AS handed_over,
      CASE
        WHEN COALESCE(a.assignments, 0) = 0 THEN 0
        ELSE ROUND(((COALESCE(c.closed_by_chatbot, 0)::numeric / a.assignments::numeric) * 100)::numeric, 1)
      END AS closed_by_chatbot_pct
    FROM days d
    LEFT JOIN assignments a ON a.day = d.day
    LEFT JOIN closed_by_chatbot c ON c.day = d.day
    LEFT JOIN handed_over h ON h.day = d.day
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

async function getBotAhtDailyRows(client, startMs, endMs) {
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
        date_trunc('day', to_timestamp(MIN("a.user_message_at") / 1000.0))::date AS day,
        MIN("a.user_message_at") AS started_at,
        MAX("Created") AS ended_at,
        (ARRAY_AGG("a.close_status" ORDER BY "Created" DESC))[1] AS latest_close_status
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
        AND "a.user_message_at" IS NOT NULL
        AND "Created" IS NOT NULL
      GROUP BY "a.ticket_id"
    ),
    filtered AS (
      SELECT
        day,
        ticket_id,
        started_at,
        ended_at,
        GREATEST(0, ended_at - started_at)::BIGINT AS duration_ms
      FROM per_ticket
      WHERE COALESCE(latest_close_status, '') <> 'NOT_CLOSED'
    ),
    joined AS (
      SELECT d.day, f.ticket_id, f.started_at, f.ended_at, f.duration_ms
      FROM days d
      LEFT JOIN filtered f ON f.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id,
      j.started_at,
      j.ended_at,
      COALESCE(j.duration_ms, 0)::BIGINT AS duration_ms
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;
    const { rows } = await client.query(sql, [startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        ticket_id: String(r.ticket_id || 'None'),
        started_at: r.started_at == null ? null : Number(r.started_at),
        ended_at: r.ended_at == null ? null : Number(r.ended_at),
        duration_ms: Number(r.duration_ms || 0),
    }));
}

async function getAgentAhtDailyRows(client, startMs, endMs) {
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
        date_trunc('day', to_timestamp(MAX("a.assigned_to_agent_at") / 1000.0))::date AS day,
        MAX("a.assigned_to_agent_at") AS started_at,
        MAX("a.agent_ended_at") AS ended_at,
        (ARRAY_AGG("a.close_status" ORDER BY "Created" DESC))[1] AS latest_close_status
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
        AND "a.assigned_to_agent_at" IS NOT NULL
        AND "a.agent_ended_at" IS NOT NULL
      GROUP BY "a.ticket_id"
    ),
    filtered AS (
      SELECT
        day,
        ticket_id,
        started_at,
        ended_at,
        GREATEST(0, ended_at - started_at)::BIGINT AS duration_ms
      FROM per_ticket
      WHERE COALESCE(latest_close_status, '') <> 'NOT_CLOSED'
    ),
    joined AS (
      SELECT d.day, f.ticket_id, f.started_at, f.ended_at, f.duration_ms
      FROM days d
      LEFT JOIN filtered f ON f.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id,
      j.started_at,
      j.ended_at,
      COALESCE(j.duration_ms, 0)::BIGINT AS duration_ms
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;
    const { rows } = await client.query(sql, [startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        ticket_id: String(r.ticket_id || 'None'),
        started_at: r.started_at == null ? null : Number(r.started_at),
        ended_at: r.ended_at == null ? null : Number(r.ended_at),
        duration_ms: Number(r.duration_ms || 0),
    }));
}

async function getTotalAhtDailyRows(client, startMs, endMs) {
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
        date_trunc('day', to_timestamp(MIN("a.user_message_at") / 1000.0))::date AS day,
        MIN("a.user_message_at") AS started_at,
        MAX("Created") AS bot_ended_at,
        MAX("a.assigned_to_agent_at") AS assigned_at,
        MAX("a.agent_ended_at") AS agent_ended_at,
        (ARRAY_AGG("a.close_status" ORDER BY "Created" DESC))[1] AS latest_close_status
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
        AND "a.user_message_at" IS NOT NULL
        AND "Created" IS NOT NULL
      GROUP BY "a.ticket_id"
    ),
    filtered AS (
      SELECT
        day,
        ticket_id,
        started_at,
        CASE
          WHEN assigned_at IS NOT NULL AND agent_ended_at IS NOT NULL THEN agent_ended_at
          ELSE bot_ended_at
        END AS ended_at
      FROM per_ticket
      WHERE COALESCE(latest_close_status, '') <> 'NOT_CLOSED'
    ),
    with_duration AS (
      SELECT
        day,
        ticket_id,
        started_at,
        ended_at,
        GREATEST(0, ended_at - started_at)::BIGINT AS duration_ms
      FROM filtered
    ),
    joined AS (
      SELECT d.day, w.ticket_id, w.started_at, w.ended_at, w.duration_ms
      FROM days d
      LEFT JOIN with_duration w ON w.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id,
      j.started_at,
      j.ended_at,
      COALESCE(j.duration_ms, 0)::BIGINT AS duration_ms
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;
    const { rows } = await client.query(sql, [startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        ticket_id: String(r.ticket_id || 'None'),
        started_at: r.started_at == null ? null : Number(r.started_at),
        ended_at: r.ended_at == null ? null : Number(r.ended_at),
        duration_ms: Number(r.duration_ms || 0),
    }));
}

async function getBotAvgResponseTimeDailyRows(client, startMs, endMs) {
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
        date_trunc('day', to_timestamp(MAX("Created") / 1000.0))::date AS day,
        COALESCE(AVG(GREATEST(0, "Created" - "a.user_message_at")), 0)::BIGINT AS avg_ms
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND "Created" IS NOT NULL
        AND "a.user_message_at" IS NOT NULL
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
      GROUP BY "a.ticket_id"
    ),
    joined AS (
      SELECT d.day, p.ticket_id, p.avg_ms
      FROM days d
      LEFT JOIN per_ticket p ON p.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id,
      COALESCE(j.avg_ms, 0)::BIGINT AS avg_ms
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;
    const { rows } = await client.query(sql, [startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        ticket_id: String(r.ticket_id || 'None'),
        avg_seconds: Math.round((Number(r.avg_ms || 0) / 1000) * 10) / 10,
    }));
}

async function getAgentFirstResponseTimeDailyRows(clientPG, startMs, endMs) {
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
        NULLIF(TRIM("a.ticket_id"), '') AS ticket_id,
        date_trunc('day', to_timestamp(MIN("a.assigned_at") / 1000.0))::date AS day,
        MIN("a.assigned_at") AS assigned_at,
        MIN("a.responded_at") AS responded_at
      FROM ${LIVE_AGENT_TABLE}
      WHERE "Created" BETWEEN $1 AND $2
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
        AND "a.assigned_at" IS NOT NULL
        AND "a.responded_at" IS NOT NULL
      GROUP BY NULLIF(TRIM("a.ticket_id"), '')
    ),
    joined AS (
      SELECT d.day, p.ticket_id, p.assigned_at, p.responded_at
      FROM days d
      LEFT JOIN per_ticket p ON p.day = d.day
    )
    SELECT
      j.day::text AS day,
      COALESCE(j.ticket_id, 'None') AS ticket_id,
      j.assigned_at,
      j.responded_at
    FROM joined j
    ORDER BY j.day ASC, (j.ticket_id = 'None') ASC, j.ticket_id ASC;
  `;
    const { rows } = await clientPG.query(sql, [startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        ticket_id: String(r.ticket_id || 'None'),
        assigned_at: r.assigned_at == null ? null : Number(r.assigned_at),
        responded_at: r.responded_at == null ? null : Number(r.responded_at),
    }));
}

async function getSlaAssignmentTimeDailyRows(client, startMs, endMs) {
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
      SELECT d.day, f.ticket_id, f.escalated_at, f.assigned_at, f.duration_ms
      FROM days d
      LEFT JOIN filtered f ON f.day = d.day
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

async function getTopInquiryTypesTicketRows24h(client, startMs, endMs) {
    const sql = `
    WITH bounds AS (
      SELECT
        date_trunc('day', (to_timestamp($2 / 1000.0) AT TIME ZONE 'Asia/Dubai'))::date AS start_day,
        date_trunc('day', (to_timestamp($3 / 1000.0) AT TIME ZONE 'Asia/Dubai'))::date AS end_day
    ),
    days AS (
      SELECT generate_series(
        (SELECT start_day FROM bounds),
        (SELECT end_day FROM bounds),
        interval '1 day'
      )::date AS day
    ),
    categories AS (
      SELECT unnest($1::text[]) AS inquiry_type
    ),
    base_rows AS (
      SELECT
        date_trunc('day', (to_timestamp("Created" / 1000.0) AT TIME ZONE 'Asia/Dubai'))::date AS day,
        to_char((to_timestamp("Created" / 1000.0) AT TIME ZONE 'Asia/Dubai'), 'YYYY-MM-DD HH24:MI:SS') || ' DXB' AS date_time,
        NULLIF(TRIM("a.inquiry_type"), '') AS inquiry_type,
        NULLIF(TRIM("a.ticket_id"), '') AS ticket_id
      FROM ${REPORT_TABLE}
      WHERE "Created" BETWEEN $2 AND $3
        AND NULLIF(TRIM("a.ticket_id"), '') IS NOT NULL
    ),
    normal_rows AS (
      SELECT day, date_time, inquiry_type, ticket_id
      FROM base_rows
      WHERE inquiry_type IS NOT NULL
        AND inquiry_type <> 'Live Agent Handover'
    ),
    live_agent_rows AS (
      SELECT day, date_time, 'Live Agent Handover'::text AS inquiry_type, ticket_id
      FROM base_rows
      WHERE inquiry_type = 'Live Agent Handover'
    ),
    combined AS (
      SELECT * FROM normal_rows
      UNION ALL
      SELECT * FROM live_agent_rows
    ),
    counts AS (
      SELECT day, inquiry_type, COUNT(*)::BIGINT AS count
      FROM combined
      GROUP BY day, inquiry_type
    ),
    expanded AS (
      SELECT
        d.day,
        c.inquiry_type,
        COALESCE(cnt.count, 0)::BIGINT AS count
      FROM days d
      CROSS JOIN categories c
      LEFT JOIN counts cnt
        ON cnt.day = d.day
       AND cnt.inquiry_type = c.inquiry_type
    ),
    ticket_rows AS (
      SELECT
        e.day,
        e.inquiry_type,
        e.count,
        c.ticket_id,
        c.date_time
      FROM expanded e
      JOIN combined c
        ON c.day = e.day
       AND c.inquiry_type = e.inquiry_type
    ),
    zero_rows AS (
      SELECT
        e.day,
        e.inquiry_type,
        e.count,
        'None'::text AS ticket_id,
        to_char((e.day::timestamp), 'YYYY-MM-DD') || ' 00:00:00 DXB' AS date_time
      FROM expanded e
      WHERE e.count = 0
    ),
    final_rows AS (
      SELECT * FROM ticket_rows
      UNION ALL
      SELECT * FROM zero_rows
    )
    SELECT
      f.day::text AS day,
      f.date_time,
      f.inquiry_type,
      f.ticket_id,
      f.count
    FROM final_rows f
    ORDER BY
      f.day ASC,
      f.count DESC,
      f.inquiry_type ASC,
      (f.ticket_id = 'None') ASC,
      f.ticket_id ASC;
  `;

    const { rows } = await client.query(sql, [INQUIRY_CATEGORIES, startMs, endMs]);
    return (rows || []).map((r) => ({
        day: String(r.day || ''),
        date_time: String(r.date_time || ''),
        inquiry_type: String(r.inquiry_type || ''),
        ticket_id: String(r.ticket_id || 'None'),
        count: Number(r.count || 0),
    }));
}

async function runExport24h(type, startMs, endMs) {
    const clientAS = await poolAS.connect();
    const clientPG = await poolPG.connect();
    try {
        switch (type) {
            case 'assignments': {
                const rows = await getAssignmentsDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Total Assignments',
                    [{ header: 'Date', key: 'day' }, { header: 'Ticket ID', key: 'ticketId' }],
                    rows.map((r) => ({ day: r.day, ticketId: r.ticket_id }))
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'closed-tickets-by-chatbot': {
                const rows = await getClosedByChatbotDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Closed by Chatbot',
                    [{ header: 'Date', key: 'day' }, { header: 'Ticket ID', key: 'ticketId' }],
                    rows.map((r) => ({ day: r.day, ticketId: r.ticket_id }))
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'handover-to-live-agent': {
                const rows = await getHandoverDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Tickets Handed Over to Live Agent',
                    [{ header: 'Date', key: 'day' }, { header: 'Ticket ID', key: 'ticketId' }],
                    rows.map((r) => ({ day: r.day, ticketId: r.ticket_id }))
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'closed-assignment-rate': {
                const rows = await getClosedAssignmentRateDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Closed Assignment Rate',
                    [
                        { header: 'Date', key: 'day' },
                        { header: 'Number of Assignments', key: 'assignments' },
                        { header: 'Closed by Chatbot', key: 'closedByChatbot' },
                        { header: 'Handed Over to Live Agent', key: 'handedOver' },
                        { header: 'Closed by Chatbot %', key: 'pct' },
                    ],
                    rows.map((r) => ({
                        day: r.day,
                        assignments: r.assignments,
                        closedByChatbot: r.closed_by_chatbot,
                        handedOver: r.handed_over,
                        pct: `${r.closed_by_chatbot_pct}%`,
                    }))
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'bot-aht': {
                const rows = await getBotAhtDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Chatbot AHT',
                    [
                        { header: 'Date', key: 'day' },
                        { header: 'Ticket ID', key: 'ticketId' },
                        { header: 'Started', key: 'started' },
                        { header: 'Ended', key: 'ended' },
                        { header: 'AHT', key: 'aht' },
                    ],
                    rows.map((r) => {
                        const ahtSeconds = r.duration_ms ? Math.round((r.duration_ms / 1000) * 10) / 10 : 0;
                        return {
                            day: r.day,
                            ticketId: r.ticket_id,
                            started: r.started_at ? formatDateTimeDubai(r.started_at) : '',
                            ended: r.ended_at ? formatDateTimeDubai(r.ended_at) : '',
                            aht: r.ticket_id === 'None' ? '0 Sec' : formatDuration(ahtSeconds),
                        };
                    })
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'agent-aht': {
                const rows = await getAgentAhtDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Agent AHT',
                    [
                        { header: 'Date', key: 'day' },
                        { header: 'Ticket ID', key: 'ticketId' },
                        { header: 'Started', key: 'started' },
                        { header: 'Ended', key: 'ended' },
                        { header: 'AHT', key: 'aht' },
                    ],
                    rows.map((r) => {
                        const ahtSeconds = r.duration_ms ? Math.round((r.duration_ms / 1000) * 10) / 10 : 0;
                        return {
                            day: r.day,
                            ticketId: r.ticket_id,
                            started: r.started_at ? formatDateTimeDubai(r.started_at) : '',
                            ended: r.ended_at ? formatDateTimeDubai(r.ended_at) : '',
                            aht: r.ticket_id === 'None' ? '0 Sec' : formatDuration(ahtSeconds),
                        };
                    })
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'total-aht': {
                const rows = await getTotalAhtDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Total AHT',
                    [
                        { header: 'Date', key: 'day' },
                        { header: 'Ticket ID', key: 'ticketId' },
                        { header: 'Started', key: 'started' },
                        { header: 'Ended', key: 'ended' },
                        { header: 'AHT', key: 'aht' },
                    ],
                    rows.map((r) => {
                        const ahtSeconds = r.duration_ms ? Math.round((r.duration_ms / 1000) * 10) / 10 : 0;
                        return {
                            day: r.day,
                            ticketId: r.ticket_id,
                            started: r.started_at ? formatDateTimeDubai(r.started_at) : '',
                            ended: r.ended_at ? formatDateTimeDubai(r.ended_at) : '',
                            aht: r.ticket_id === 'None' ? '0 Sec' : formatDuration(ahtSeconds),
                        };
                    })
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'bot-avg-response-time': {
                const rows = await getBotAvgResponseTimeDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Chatbot Avg Response Time',
                    [
                        { header: 'Date', key: 'day' },
                        { header: 'Ticket ID', key: 'ticketId' },
                        { header: 'Average Bot Response Time', key: 'avg' },
                    ],
                    rows.map((r) => ({
                        day: r.day,
                        ticketId: r.ticket_id,
                        avg: r.ticket_id === 'None' ? '0 Sec' : formatDuration(r.avg_seconds),
                    }))
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'agent-first-response-time': {
                const rows = await getAgentFirstResponseTimeDailyRows(clientPG, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'Agent First Response Time',
                    [
                        { header: 'Date', key: 'day' },
                        { header: 'Ticket ID', key: 'ticketId' },
                        { header: 'Assigned Time', key: 'assigned' },
                        { header: 'Responded Time', key: 'responded' },
                    ],
                    rows.map((r) => ({
                        day: r.day,
                        ticketId: r.ticket_id,
                        assigned: r.assigned_at ? formatDateTimeDubai(r.assigned_at) : '',
                        responded: r.responded_at ? formatDateTimeDubai(r.responded_at) : '',
                    }))
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'sla-assignment-time': {
                const rows = await getSlaAssignmentTimeDailyRows(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
                    'SLA Assignment Time',
                    [
                        { header: 'Date', key: 'day' },
                        { header: 'Ticket ID', key: 'ticketId' },
                        { header: 'Escalation Time', key: 'esc' },
                        { header: 'Agent Assignment Time', key: 'asg' },
                        { header: 'SLA', key: 'sla' },
                    ],
                    rows.map((r) => {
                        const seconds = r.duration_ms ? Math.round((r.duration_ms / 1000) * 10) / 10 : 0;
                        return {
                            day: r.day,
                            ticketId: r.ticket_id,
                            esc: r.escalated_at ? formatDateTimeDubai(r.escalated_at) : '',
                            asg: r.assigned_at ? formatDateTimeDubai(r.assigned_at) : '',
                            sla: r.ticket_id === 'None' ? '0 Sec' : formatDuration(seconds),
                        };
                    })
                );
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            case 'top-inquiry-types': {
                const rows = await getTopInquiryTypesTicketRows24h(clientAS, startMs, endMs);
                const wb = new ExcelJS.Workbook();
                addSheetFromRows(
                    wb,
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
                return workbookBufferToResponse(wb, `${type}.xlsx`);
            }

            default:
                return json(400, { error: `Unsupported type: ${type}` });
        }
    } finally {
        clientAS.release();
        clientPG.release();
    }
}

async function getRealtime24h(type) {
    if (!DB_HOST || !DB_USER || !DB_PASS || !REPORT_TABLE) {
        return json(500, { error: 'Missing required env vars: DB_HOST, userName, DB_PASSWORD, REPORT_TABLE' });
    }

    const { startMs, endMs } = last24HoursRangeDubai();
    return await runExport24h(type, startMs, endMs);
}
module.exports = { getRealtime24h };
exports.handler = async (event) => {
    try {
        if (event?.requestContext?.http?.method === 'OPTIONS' || event?.httpMethod === 'OPTIONS') {
            return { statusCode: 200, headers: corsHeaders, body: '' };
        }

        const qs = event.queryStringParameters || event.query || {};
        const type = String(qs.type || '').trim();
        if (!type) return json(400, { error: 'Missing query param: type' });

        return await getRealtime24h(type);
    } catch (err) {
        console.error('reporting-export last24h error:', err);
        return json(500, { error: err?.message || 'Internal server error' });
    }
};