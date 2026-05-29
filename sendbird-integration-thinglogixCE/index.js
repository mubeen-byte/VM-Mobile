var AWS = require("aws-sdk");
const lambda = new AWS.Lambda();
const axios = require("axios");

const { Client } = require("pg");
const env = process.env.ENV ? process.env.ENV : "prod";
const APP_ID = process.env.APP_ID;
// const apiToken = process.env.DESK_API_TOKEN;
const deskToken = process.env.DESK_API_TOKEN;
const API_KEY = process.env.API_KEY;

const DB_HOST = process.env.DB_HOST;
const DB_USER = process.env.userName;
const DB_PASS = process.env.DB_PASSWORD;
const DB_NAME = process.env.DB_NAME || "AdvancedSearch";
const DB_PORT = parseInt(process.env.DB_PORT || "5432", 10);
const LLM_TABLE = process.env.LLM_TABLE;
const LIVE_AGENT_DB_NAME = "postgres";
const LIVE_AGENT_TABLE = process.env.LIVE_AGENT_TABLE;

const REOPEN_DB_NAME = "postgres";
const REOPEN_TABLE = 'public.ticket_reopen_tracking';

exports.handler = async (event, context, callback) => {
  event = event.body ? event.body : event;
  if (typeof event === "string") {
    try { event = JSON.parse(event); } catch (e) {}
  }
  if (event && typeof event.body === "string") {
    try { event.body = JSON.parse(event.body); } catch (e) {}
  }
  console.log("Event: ", JSON.stringify(event));
  if(event?.eventType == 'TICKET.CREATED') {
    console.log("create ticket event");
    const response = addMemberInGroup(event);
    return callback(null,response);
  } else if (event.eventType === "MESSAGE.RECEIVED" && event?.botWebhookEvent ) {
  console.log("Processing bot MESSAGE.RECEIVED event...");

  const data = event.ticket;
  const chatMessage = event.botWebhookEvent.chatMessage;
  const rawLastMessageAt =
      chatMessage?.createdAt ||
      data?.lastMessageAt ||
      data?.lastUserMessageAt ||
      null;
  const lastMessageAt = rawLastMessageAt ? Date.parse(rawLastMessageAt) : null;
  const closeStatus =
      data?.closeStatus ||
      null;

  try {
    const payload = {
      channelUrl: data.channelUrl,
      customer: data.customer,
      ticketId: String(data.id),
      sessionId: `sendbird:${data.customer?.id}`,
      message: chatMessage.message,
      groupId: String(data.group.id),
      botWebhookEventId: String(event.botWebhookEvent.id),
      sendbirdBotId: String(event.botWebhookEvent.bot),
      lastMessageAt,
      closeStatus
    };
    console.log("Payload for MESSAGE.RECEIVED:", JSON.stringify(payload));

    const lambdaPayload = {
      sessionId: String(payload.sessionId || "default-session"),
      message: payload.message,
      ticketId: payload.ticketId,
      groupId: payload.groupId,
      channelUrl: payload.channelUrl,
      botWebhookEventId: payload.botWebhookEventId,
      sendbirdBotId: payload.sendbirdBotId,
      lastMessageAt: payload.lastMessageAt,
      closeStatus: payload.closeStatus,
      sendbird_user_name: payload.customer?.displayName,
      sendbird_product_type: data?.customFields[0]?.value ||'default',
      chatbots: JSON.parse(process.env.chatbots),
      sessionDetails: {
        intentName: null,
        messageFormat: "PlainText",
        sessionAttributes: {
          channelUrl: payload.channelUrl,
          channel: "sendbird",
          sessionId: payload.sessionId
        },
        multiLanguage: true,
        callCustomFunctions: true,
        language: null,
        slotToElicit: null,
        slots: {},
      },
    };

    console.log("Lambda payload for bot validation:", JSON.stringify(lambdaPayload));
    await invokeValidateBotLambda(lambdaPayload);

    return callback(null, { message: "MESSAGE.RECEIVED processed successfully." });
  } catch (error) {
    console.error("Error processing event:", error);
    return callback(error);
  }
}
  // } else if(event?.type =='MESG' && event?.category == 'group_channel:message_send' && event?.sdk == 'Android' && event?.sender?.user_id) {
  //   console.log("Processing chat MESSAGE event...");
  
  //   try {
  //     const lambdaPayload = {
  //       sessionId: `sendbird:${event?.sender?.user_id}`.replace("@",""),
  //       message: event.payload.message,
  //       chatbots: JSON.parse(process.env.chatbots),
  //       sessionDetails: {
  //         intentName: null,
  //         messageFormat: "PlainText",
  //         sessionAttributes: {
  //           channelUrl: event.channel.channel_url,
  //           channel: "sendbird",
  //           userId: event?.sender?.user_id
  //         },
  //         multiLanguage: true,
  //         callCustomFunctions: true,
  //         language: null,
  //         slotToElicit: null,
  //         slots: {},
  //       },
  //     };
  
  //     console.log("Lambda payload for bot validation:", JSON.stringify(lambdaPayload));
  //     await invokeValidateBotLambda(lambdaPayload);
  
  //     return callback(null, { message: "chat platform processed successfully." });
  //   } catch (error) {
  //     console.error("Error processing event:", error);
  //     return callback(error);
  //   }
  // }
  else if (event?.eventType === "TICKET.CLOSED" || event?.eventType === "TICKET.STATUS.UPDATED") {
    console.log(`Processing ${event.eventType} event...`);

    const t = event.data;
    const ticketId = t?.id != null ? String(t.id) : null;
    const closeStatus = t?.closeStatus || null;

    if (!ticketId || !closeStatus) {
      console.log("Missing ticketId or closeStatus; skipping update.", { ticketId, closeStatus });
      return callback(null, { message: "Missing ticketId/closeStatus. No update performed." });
    }

    try {
      const incrementReopenCount = event?.eventType === "TICKET.STATUS.UPDATED";
      const prev = incrementReopenCount ? await getLatestTicketRow(ticketId) : null;

      await updateLatestTicketCloseStatus(ticketId, closeStatus, incrementReopenCount);

      if (incrementReopenCount && prev && String(prev.close_status || '') !== 'NOT_CLOSED' && closeStatus === 'NOT_CLOSED') {
        const closedAt = prev.updated != null ? Number(prev.updated) : (prev.created != null ? Number(prev.created) : null);
        const reopenedAt = toMs(t?.updatedAt) || toMs(event?.createdAt) || Date.now();
        await insertReopenTracking(ticketId, closedAt, reopenedAt);
      }

      const raAny = t?.recentAssignment || null;
      const agentAny = raAny?.agent || null;
      const agentEmailAny = agentAny?.email != null ? String(agentAny.email).trim() : "";

      if (agentEmailAny) {
        const assignedAtRawAny = raAny?.assignedAt || null;
        const endedAtRawAny = raAny?.endedAt || null;

        const assignedAtMsAny = assignedAtRawAny ? Date.parse(assignedAtRawAny) : null;
        const endedAtMsAny = endedAtRawAny ? Date.parse(endedAtRawAny) : null;

        await updateLatestTicketAgentTimes(
          ticketId,
          assignedAtMsAny && !Number.isNaN(assignedAtMsAny) ? assignedAtMsAny : null,
          endedAtMsAny && !Number.isNaN(endedAtMsAny) ? endedAtMsAny : null
        );
      }
    
      if (event?.eventType === "TICKET.STATUS.UPDATED") {
        const ra = t?.recentAssignment || null;
        const agent = ra?.agent || null;    
        const agentEmail = agent?.email != null ? String(agent.email) : "";
        const respondedAtRaw = ra?.respondedAt || ra?.responsedAt || null;
        const ESCALATION_MSG =
          "Alright, I'm now connecting you to a live agent. Please wait a moment while I transfer you...";
        const lastMessageRaw =
          event?.data?.lastMessage ??
          event?.data?.ticket?.lastMessage ??
          t?.lastMessage ??
          "";
    
        const lastMessage = String(lastMessageRaw || "").trim();
        const hasAgentEmail = !!(agentEmail && agentEmail.trim() !== "");
        const shouldEscalate = hasAgentEmail || lastMessage === ESCALATION_MSG;
    
        if (shouldEscalate) {
          await updateLatestTicketEscalated(ticketId, true, Date.now());
        }
    
        if (hasAgentEmail && respondedAtRaw && String(respondedAtRaw).trim() !== "") {
          const assignedAtRaw = ra?.assignedAt || null;
    
          const assignedAtMs = assignedAtRaw ? Date.parse(assignedAtRaw) : null;
          const respondedAtMs = respondedAtRaw ? Date.parse(respondedAtRaw) : null;
    
          if (respondedAtMs && !Number.isNaN(respondedAtMs)) {
            await insertLiveAgentRecord({
              ticketId,
              agentId: agent?.id != null ? Number(agent.id) : null,
              agentName: agent?.displayName != null ? String(agent.displayName) : null,
              agentEmail: agentEmail.trim(),
              assignedAt: assignedAtMs && !Number.isNaN(assignedAtMs) ? assignedAtMs : null,
              respondedAt: respondedAtMs,
            });
          } else {
            console.log("Skipping live_agent_data insert: respondedAt could not be parsed to ms.", { respondedAtRaw });
          }
        } else {
          console.log("Skipping live_agent_data insert: conditions not met.", {
            hasEmail: !!(agentEmail && agentEmail.trim() !== ""),
            respondedAtRaw,
          });
        }
      }
    
      return callback(null, { message: `${event.eventType} processed successfully.` });
    } catch (error) {
      console.error("Error processing event:", error);
      return callback(error);
    }
  }
  else {
    console.log("Unhandled event or structure.");
    return callback(null, { message: "Unhandled event type or structure." });
  }
};

async function invokeValidateBotLambda(payload) {
  const params = {
    FunctionName: `validate-bot-request-thinglogixCE-${env}`,
    Payload: JSON.stringify({ body: payload }),
    InvocationType: "Event",
  };

  try {
    const data = await lambda.invoke(params).promise();
    console.log("Lambda invoked successfully:", JSON.stringify(data));
  } catch (err) {
    console.error("Error invoking Lambda:", err);
    throw err;
  }
}

async function addMemberInGroup (event) {
  const channelUrl = event?.data?.channelUrl;
    const payload = {
      "user_ids" : [process.env.SENDBIRD_BOT_USER_ID]
    };
    console.log("add member payload =>",JSON.stringify(payload));

    try {
        const response = await axios.post(
            `https://api-${APP_ID}.sendbird.com/v3/group_channels/${channelUrl}/invite`,
            payload,
            {
                headers: {
                    'Content-Type': 'application/json',
                    'Api-Token': API_KEY,
                }
            }
        );

        console.log("Member added successfully:", response?.data);
        return response?.data;
        
    } catch (error) {
        console.error("Error sending message:", error);
        throw error;
    }
}

async function updateLatestTicketCloseStatus(ticketId, closeStatus, incrementReopenCount) {
  const client = new Client({
    host: DB_HOST,
    user: DB_USER,
    password: DB_PASS,
    database: DB_NAME,
    port: DB_PORT,
    ssl: process.env.DB_SSL === "true" ? { rejectUnauthorized: false } : undefined,
  });

  const updatedMs = Date.now();

  const sql = `
    UPDATE ${LLM_TABLE} t
    SET "a.close_status" = $1,
        "Updated" = $2,
        "a.reopen_count" = CASE
          WHEN $4::boolean IS TRUE
               AND COALESCE(t."a.close_status",'') <> 'NOT_CLOSED'
               AND $1 = 'NOT_CLOSED'
          THEN COALESCE(t."a.reopen_count", 0) + 1
          ELSE COALESCE(t."a.reopen_count", 0)
        END
    WHERE "deviceId" = (
      SELECT "deviceId"
      FROM ${LLM_TABLE}
      WHERE "a.ticket_id" = $3
      ORDER BY "Created" DESC
      LIMIT 1
    )
    RETURNING "deviceId", "a.ticket_id", "a.close_status", "a.reopen_count", "Created", "Updated";
  `;

  await client.connect();
  try {
    const res = await client.query(sql, [closeStatus, updatedMs, ticketId, !!incrementReopenCount]);
    console.log("Updated rows:", res.rowCount, res.rows?.[0] || null);
    return res.rows?.[0] || null;
  } finally {
    await client.end();
  }
}

async function updateLatestTicketAgentTimes(ticketId, assignedAtMs, endedAtMs) {
  if (!ticketId) return;

  const hasAssigned = assignedAtMs != null && Number.isFinite(assignedAtMs);
  const hasEnded = endedAtMs != null && Number.isFinite(endedAtMs);
  if (!hasAssigned && !hasEnded) return;

  const client = new Client({
    host: DB_HOST,
    user: DB_USER,
    password: DB_PASS,
    database: DB_NAME,
    port: DB_PORT,
    ssl: process.env.DB_SSL === "true" ? { rejectUnauthorized: false } : undefined,
  });

  const nowMs = Date.now();

  const sql = `
  UPDATE ${LLM_TABLE} t
  SET
    "a.assigned_to_agent_at" = CASE
      WHEN t."a.assigned_to_agent_at" IS NULL THEN $2
      ELSE t."a.assigned_to_agent_at"
    END,
    "a.agent_ended_at" = COALESCE($3, t."a.agent_ended_at"),
    "Updated" = GREATEST(COALESCE(t."Updated", 0), $4)
  WHERE "deviceId" = (
    SELECT "deviceId"
    FROM ${LLM_TABLE}
    WHERE "a.ticket_id" = $1
    ORDER BY "Created" DESC
    LIMIT 1
  )
  RETURNING "deviceId","a.ticket_id","a.assigned_to_agent_at","a.agent_ended_at","Created","Updated";
`;

  try {
    await client.connect();
    const res = await client.query(sql, [
      ticketId,
      hasAssigned ? Math.floor(assignedAtMs) : null,
      hasEnded ? Math.floor(endedAtMs) : null,
      nowMs,
    ]);
    console.log("updateLatestTicketAgentTimes updated rows:", res.rowCount, res.rows?.[0] || null);
    return res.rows?.[0] || null;
  } finally {
    await client.end().catch(() => {});
  }
}

async function updateLatestTicketEscalated(ticketId, escalated, escalatedAtMs) {
  if (!ticketId) return;
  if (escalated !== true) return;

  const client = new Client({
    host: DB_HOST,
    user: DB_USER,
    password: DB_PASS,
    database: DB_NAME,
    port: DB_PORT,
    ssl: process.env.DB_SSL === "true" ? { rejectUnauthorized: false } : undefined,
  });

  const nowMs = Date.now();
  const updatedMs = (escalatedAtMs && Number.isFinite(escalatedAtMs)) ? Math.floor(escalatedAtMs) : nowMs;

  const sql = `
    UPDATE ${LLM_TABLE} t
    SET
      "a.escalated" = TRUE,
      "Updated" = GREATEST(COALESCE(t."Updated", 0), $2)
    WHERE "deviceId" = (
      SELECT "deviceId"
      FROM ${LLM_TABLE}
      WHERE "a.ticket_id" = $1
      ORDER BY "Created" DESC
      LIMIT 1
    )
    RETURNING "deviceId","a.ticket_id","a.escalated","a.escalated_at","Created","Updated";
  `;

  try {
    await client.connect();
    const res = await client.query(sql, [ticketId, updatedMs]);
    console.log("updateLatestTicketEscalated updated rows:", res.rowCount, res.rows?.[0] || null, { ticketId });
    return res.rows?.[0] || null;
  } catch (e) {
    console.error("updateLatestTicketEscalated failed:", e?.message || e, { ticketId });
    throw e;
  } finally {
    await client.end().catch(() => {});
  }
}

async function insertLiveAgentRecord({ ticketId, agentId, agentName, agentEmail, assignedAt, respondedAt }) {
  const client = new Client({
    host: DB_HOST,
    user: DB_USER,
    password: DB_PASS,
    database: LIVE_AGENT_DB_NAME,
    port: DB_PORT,
    ssl: process.env.DB_SSL === "true" ? { rejectUnauthorized: false } : undefined,
  });

  const nowMs = Date.now();

  const sql = `
  INSERT INTO ${LIVE_AGENT_TABLE}
    ("a.ticket_id","a.agent_id","a.agent_name","a.agent_email","Created","Updated","a.assigned_at","a.responded_at")
  SELECT
    $1,$2,$3,$4,$5,$6,$7,$8
  WHERE NOT EXISTS (
    SELECT 1
    FROM ${LIVE_AGENT_TABLE} x
    WHERE x."a.ticket_id" = $1
      AND COALESCE(x."a.assigned_at", -1::BIGINT) = COALESCE($7::BIGINT, -1::BIGINT)
      AND COALESCE(x."a.responded_at", -1::BIGINT) = COALESCE($8::BIGINT, -1::BIGINT)
  )
  RETURNING "id","a.ticket_id","a.agent_id","a.agent_email","a.assigned_at","a.responded_at","Created","Updated";
`;

  await client.connect();
  try {
    const res = await client.query(sql, [
      ticketId,
      agentId,
      agentName,
      agentEmail,
      nowMs,
      nowMs,
      assignedAt,
      respondedAt,
    ]);

    console.log("live_agent_data insert rows:", res.rowCount, res.rows?.[0] || null);
    return res.rows?.[0] || null;
  } finally {
    await client.end();
  }
}

function toMs(v) {
  const ms = v ? Date.parse(String(v)) : NaN;
  return Number.isFinite(ms) ? ms : null;
}

async function getLatestTicketRow(ticketId) {
  const client = new Client({
    host: DB_HOST,
    user: DB_USER,
    password: DB_PASS,
    database: DB_NAME,
    port: DB_PORT,
    ssl: process.env.DB_SSL === "true" ? { rejectUnauthorized: false } : undefined,
  });

  const sql = `
    SELECT "a.close_status" AS close_status, "Created" AS created, "Updated" AS updated
    FROM ${LLM_TABLE}
    WHERE "a.ticket_id" = $1
    ORDER BY "Created" DESC
    LIMIT 1;
  `;

  await client.connect();
  try {
    const res = await client.query(sql, [ticketId]);
    return res.rows?.[0] || null;
  } finally {
    await client.end().catch(() => {});
  }
}

async function insertReopenTracking(ticketId, closedAt, reopenedAt) {
  const client = new Client({
    host: DB_HOST,
    user: DB_USER,
    password: DB_PASS,
    database: REOPEN_DB_NAME,
    port: DB_PORT,
    ssl: process.env.DB_SSL === "true" ? { rejectUnauthorized: false } : undefined,
  });

  const nowMs = Date.now();
  const sql = `
    INSERT INTO ${REOPEN_TABLE}
      ("a.ticket_id","a.ticket_closed_at","a.reopened_at","Created","Updated")
    VALUES
      ($1,$2,$3,$4,$5);
  `;

  await client.connect();
  try {
    await client.query(sql, [ticketId, closedAt, reopenedAt, nowMs, nowMs]);
  } finally {
    await client.end().catch(() => {});
  }
}