'use strict';

const { getFileFromS3 } = require('./last12mo');
const { getRealtime24h } = require('./last24h');

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

function parseType(raw) {
  const t = String(raw || '').trim();
  if (!t) return { mode: '', baseType: '' };

  const lower = t.toLowerCase();

  const is12 =
    lower.includes('12-mo') ||
    lower.includes('12mo') ||
    lower.includes('12_month') ||
    lower.includes('12month') ||
    lower.includes('12-month');

  const is24 =
    lower.includes('24-h') ||
    lower.includes('24h') ||
    lower.includes('24_hour') ||
    lower.includes('24hour') ||
    lower.includes('24-hour');

  const baseType = lower
    .replace(/(^|[-_])12[-_]?mo(nths)?($|[-_])/g, '$1')
    .replace(/(^|[-_])12[-_]?month(s)?($|[-_])/g, '$1')
    .replace(/(^|[-_])24[-_]?h(ours)?($|[-_])/g, '$1')
    .replace(/(^|[-_])24[-_]?hour(s)?($|[-_])/g, '$1')
    .replace(/[-_]+$/g, '')
    .replace(/^[-_]+/g, '')
    .trim();

  const mode = is12 ? '12mo' : is24 ? '24h' : '';
  return { mode, baseType: baseType || lower };
}

exports.handler = async (event) => {
  try {
    if (event?.requestContext?.http?.method === 'OPTIONS' || event?.httpMethod === 'OPTIONS') {
      return { statusCode: 200, headers: corsHeaders, body: '' };
    }

    const qs = event.queryStringParameters || event.query || {};
    const rawType = qs.type;

    if (!rawType) return json(400, { error: 'Missing query param: type' });

    const { mode, baseType } = parseType(rawType);
    if (!mode) return json(400, { error: 'type must include either 12-mo or 24-h indicator' });
    if (!baseType) return json(400, { error: 'Invalid type' });

    if (mode === '12mo') return await getFileFromS3(baseType);
    return await getRealtime24h(baseType);
  } catch (err) {
    console.error('reporting-export router error:', err);
    return {
      statusCode: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err?.message || 'Internal server error' }),
    };
  }
};