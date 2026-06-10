import {
  GRAPH_BASE,
  DELETE_REVIEW_FOLDER,
  MAX_MESSAGES,
} from "./config.js";
import { getToken } from "./auth.js";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Escape a single-quoted OData string literal (doubles any apostrophes).
function odataLiteral(value) {
  return value.replace(/'/g, "''");
}

// Core fetch wrapper: adds the bearer token, handles top-level 429 throttling,
// and parses JSON. `absolute` = true when following an @odata.nextLink URL.
async function graph(path, options = {}, absolute = false) {
  const token = await getToken();
  if (!token) return null; // page is redirecting for re-auth
  const url = absolute ? path : `${GRAPH_BASE}${path}`;

  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (res.status === 429) {
    const retryAfter = parseInt(res.headers.get("Retry-After") || "5", 10);
    await sleep((retryAfter || 5) * 1000);
    return graph(path, options, absolute);
  }

  if (!res.ok) {
    let detail = "";
    try {
      detail = JSON.stringify((await res.json()).error);
    } catch {
      detail = await res.text();
    }
    throw new Error(`Graph ${res.status}: ${detail}`);
  }

  if (res.status === 204) return null;
  return res.json();
}

export async function getProfile() {
  return graph("/me?$select=displayName,userPrincipalName,mail");
}

// Find the "Delete Review" folder, creating it at the top level if missing.
// Returns the folder id.
export async function ensureDeleteReviewFolder() {
  const filter = encodeURIComponent(
    `displayName eq '${odataLiteral(DELETE_REVIEW_FOLDER)}'`
  );
  const existing = await graph(`/me/mailFolders?$filter=${filter}&$top=10`);
  if (existing && existing.value && existing.value.length > 0) {
    return existing.value[0].id;
  }
  const created = await graph("/me/mailFolders", {
    method: "POST",
    body: JSON.stringify({ displayName: DELETE_REVIEW_FOLDER }),
  });
  return created.id;
}

// Find all messages from a sender. scope = "inbox" (default) or "all" folders.
// onProgress(countSoFar) is called as pages stream in.
export async function findMessagesFromSender(sender, scope, onProgress) {
  const filter = encodeURIComponent(
    `from/emailAddress/address eq '${odataLiteral(sender)}'`
  );
  const base =
    scope === "all" ? "/me/messages" : "/me/mailFolders/inbox/messages";
  let path = `${base}?$filter=${filter}&$select=id,subject,from,receivedDateTime&$top=100`;

  const messages = [];
  let data = await graph(path);
  while (data) {
    for (const m of data.value || []) messages.push(m);
    if (onProgress) onProgress(messages.length);
    if (messages.length >= MAX_MESSAGES) break;
    const next = data["@odata.nextLink"];
    if (!next) break;
    data = await graph(next, {}, true);
  }
  return messages;
}

// Move messages into the destination folder using Graph JSON batching
// (20 requests per round trip). Retries throttled/failed sub-requests.
// onProgress(done, total) fires after each batch. Returns { moved, failed }.
export async function moveMessages(ids, destinationId, onProgress) {
  let moved = 0;
  const failedIds = [];

  for (let i = 0; i < ids.length; i += 20) {
    const chunk = ids.slice(i, i + 20);
    const result = await runMoveBatch(chunk, destinationId);
    moved += result.ok.length;
    failedIds.push(...result.failed);
    if (onProgress) onProgress(Math.min(i + 20, ids.length), ids.length);
  }

  // One retry pass for anything that failed (often transient throttling).
  const stillFailed = [];
  if (failedIds.length) {
    await sleep(2000);
    for (let i = 0; i < failedIds.length; i += 20) {
      const chunk = failedIds.slice(i, i + 20);
      const result = await runMoveBatch(chunk, destinationId);
      moved += result.ok.length;
      stillFailed.push(...result.failed);
    }
  }

  return { moved, failed: stillFailed };
}

async function runMoveBatch(ids, destinationId) {
  const body = {
    requests: ids.map((id, idx) => ({
      id: String(idx),
      method: "POST",
      url: `/me/messages/${id}/move`,
      headers: { "Content-Type": "application/json" },
      body: { destinationId },
    })),
  };

  const resp = await graph("/$batch", {
    method: "POST",
    body: JSON.stringify(body),
  });

  const ok = [];
  const failed = [];
  const byIndex = new Map(ids.map((id, idx) => [String(idx), id]));
  for (const r of (resp && resp.responses) || []) {
    const id = byIndex.get(r.id);
    if (r.status >= 200 && r.status < 300) ok.push(id);
    else failed.push(id);
  }
  return { ok, failed };
}
