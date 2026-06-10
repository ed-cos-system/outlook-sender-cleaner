import {
  initAuth,
  login,
  logout,
  getActiveAccount,
} from "./auth.js";
import {
  getProfile,
  ensureDeleteReviewFolder,
  findMessagesFromSender,
  moveMessages,
  scanRecentMessages,
} from "./graph.js";
import { rankSenders } from "./senders.js";
import { CLIENT_ID } from "./config.js";

const $ = (id) => document.getElementById(id);
const els = {};
[
  "account", "account-name", "signout", "signin-card", "signin", "app-card",
  "sender", "find", "results", "count", "count-sender", "preview", "move",
  "progress-wrap", "bar-fill", "progress-text", "status",
  "scan-window", "scan", "scan-summary", "sender-list", "sweep",
].forEach((id) => (els[id] = $(id)));

let matched = []; // currently found messages awaiting a move (manual flow)
let ranked = []; // ranked sender candidates from the last scan (picker flow)

function setStatus(msg, kind = "info") {
  els.status.textContent = msg;
  els.status.dataset.kind = kind;
}

function isEmail(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
}

// Accept whatever Outlook shows, e.g.
//   ProjectManagement.com Newsletter <newsletter@email.projectmanagement.com>
// and pull out just the address. Falls back to the trimmed input when there
// are no angle brackets (a plain address pasted on its own).
function extractEmail(raw) {
  if (!raw) return "";
  const v = raw.trim();
  const bracketed = v.match(/<([^>]+)>/);
  const candidate = bracketed ? bracketed[1] : v;
  return candidate.trim().replace(/^["']+|["']+$/g, "").toLowerCase();
}

function showSignedIn(account) {
  els["signin-card"].hidden = true;
  els["app-card"].hidden = false;
  els.account.hidden = false;
  els["account-name"].textContent =
    account.username || account.name || "signed in";
}

function showSignedOut() {
  els["signin-card"].hidden = false;
  els["app-card"].hidden = true;
  els.account.hidden = true;
}

function resetResults() {
  matched = [];
  els.results.hidden = true;
  els.preview.innerHTML = "";
  els.move.disabled = true;
  els["progress-wrap"].hidden = true;
  els["bar-fill"].style.width = "0%";
}

async function onFind() {
  const sender = extractEmail(els.sender.value);
  if (!isEmail(sender)) {
    setStatus(
      "Couldn't find an email address in that. Paste the sender from Outlook, or just the address.",
      "error"
    );
    return;
  }
  // Show the cleaned address so it's clear what we're searching for.
  els.sender.value = sender;
  resetResults();
  els.find.disabled = true;
  setStatus(`Searching for mail from ${sender}…`);

  try {
    const scope =
      document.querySelector('input[name="scope"]:checked').value;
    matched = await findMessagesFromSender(sender, scope, (n) =>
      setStatus(`Found ${n} so far…`)
    );

    els.count.textContent = String(matched.length);
    els["count-sender"].textContent = sender;
    renderPreview(matched);
    els.results.hidden = false;
    els.move.disabled = matched.length === 0;

    if (matched.length === 0) {
      setStatus(`No messages from ${sender} in the selected scope.`, "info");
    } else {
      setStatus(
        `Found ${matched.length}. Review below, then move them.`,
        "info"
      );
    }
  } catch (err) {
    setStatus(err.message || String(err), "error");
  } finally {
    els.find.disabled = false;
  }
}

function renderPreview(messages) {
  els.preview.innerHTML = "";
  messages.slice(0, 12).forEach((m) => {
    const li = document.createElement("li");
    const subj = document.createElement("span");
    subj.className = "subj";
    subj.textContent = m.subject || "(no subject)";
    const date = document.createElement("span");
    date.className = "mono date";
    date.textContent = m.receivedDateTime
      ? new Date(m.receivedDateTime).toLocaleDateString()
      : "";
    li.append(subj, date);
    els.preview.appendChild(li);
  });
  if (messages.length > 12) {
    const li = document.createElement("li");
    li.className = "more";
    li.textContent = `+ ${messages.length - 12} more`;
    els.preview.appendChild(li);
  }
}

// --- Sender picker ---------------------------------------------------------

function resetPicker() {
  ranked = [];
  els["sender-list"].innerHTML = "";
  els["sender-list"].hidden = true;
  els.sweep.hidden = true;
  els.sweep.disabled = true;
  els["scan-summary"].textContent = "";
}

function cutoffISOForMonths(months) {
  const d = new Date();
  d.setMonth(d.getMonth() - months);
  return d.toISOString();
}

function selectedAddresses() {
  return [...els["sender-list"].querySelectorAll(".sender-check:checked")].map(
    (c) => c.dataset.address
  );
}

function updateSweepButton() {
  const n = selectedAddresses().length;
  els.sweep.disabled = n === 0;
  els.sweep.textContent = n
    ? `Move ${n} sender${n === 1 ? "" : "s"} to Delete Review`
    : "Move selected to Delete Review";
}

function renderSenderList(list) {
  const ul = els["sender-list"];
  ul.innerHTML = "";
  list.forEach((s) => {
    const li = document.createElement("li");
    li.className = "sender-row";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "sender-check";
    cb.dataset.address = s.address;
    cb.addEventListener("change", updateSweepButton);

    const main = document.createElement("div");
    main.className = "sender-main";

    const top = document.createElement("div");
    top.className = "sender-top";
    const nm = document.createElement("span");
    nm.className = "sender-name";
    nm.textContent = s.name || s.address;
    const addr = document.createElement("span");
    addr.className = "sender-addr mono";
    addr.textContent = s.address;
    top.append(nm, addr);

    const meta = document.createElement("div");
    meta.className = "sender-meta";
    const cnt = document.createElement("span");
    cnt.className = "sender-count";
    cnt.textContent = `${s.count}×`;
    const date = document.createElement("span");
    date.className = "sender-date mono";
    date.textContent = s.latest
      ? new Date(s.latest).toLocaleDateString()
      : "";
    meta.append(cnt, date);
    for (const t of s.tags) {
      const tag = document.createElement("span");
      tag.className = "tag";
      tag.textContent = t;
      meta.appendChild(tag);
    }

    const subj = document.createElement("div");
    subj.className = "sender-subj";
    subj.textContent = s.sampleSubject;

    main.append(top, meta, subj);
    // Clicking anywhere in the row body toggles its checkbox.
    main.addEventListener("click", () => {
      cb.checked = !cb.checked;
      updateSweepButton();
    });

    li.append(cb, main);
    ul.appendChild(li);
  });
}

async function onScan() {
  const months = parseInt(els["scan-window"].value, 10) || 3;
  const cutoff = cutoffISOForMonths(months);
  const label = `${months} month${months === 1 ? "" : "s"}`;

  resetPicker();
  resetResults(); // clear any manual-flow results to avoid confusion
  els.scan.disabled = true;
  setStatus(`Scanning your inbox (last ${label})…`);

  try {
    const messages = await scanRecentMessages(cutoff, (n) =>
      setStatus(`Scanned ${n} messages…`)
    );
    ranked = rankSenders(messages);
    renderSenderList(ranked);
    els["scan-summary"].textContent = `${messages.length} messages · ${ranked.length} senders · last ${label}`;
    els["sender-list"].hidden = ranked.length === 0;
    els.sweep.hidden = ranked.length === 0;
    updateSweepButton();

    if (ranked.length === 0) {
      setStatus(`No mail found in the last ${label}.`, "info");
    } else {
      setStatus(
        `Found ${ranked.length} senders. Tick the ones to clear, then move them.`,
        "info"
      );
    }
  } catch (err) {
    setStatus(err.message || String(err), "error");
  } finally {
    els.scan.disabled = false;
  }
}

async function onSweep() {
  const addresses = selectedAddresses();
  if (addresses.length === 0) return;
  const scope = document.querySelector('input[name="scope"]:checked').value;

  els.sweep.disabled = true;
  els.scan.disabled = true;
  els.find.disabled = true;
  els["progress-wrap"].hidden = false;
  els["bar-fill"].style.width = "0%";
  setStatus(
    `Gathering mail from ${addresses.length} sender${addresses.length === 1 ? "" : "s"}…`
  );

  try {
    // Collect every message id for each selected sender — all of their mail in
    // the chosen scope, not just what the recent scan saw. This honors "the
    // sender I pick gets ALL of their mail moved."
    const allIds = [];
    for (let i = 0; i < addresses.length; i++) {
      const addr = addresses[i];
      els["progress-text"].textContent = `sender ${i + 1} / ${addresses.length}`;
      const msgs = await findMessagesFromSender(addr, scope, (n) =>
        setStatus(`Gathering ${addr}: ${n}…`)
      );
      for (const m of msgs) allIds.push(m.id);
    }

    if (allIds.length === 0) {
      setStatus("Nothing to move for the selected senders.", "info");
      els["progress-wrap"].hidden = true;
      els.sweep.disabled = false;
      return;
    }

    const folderId = await ensureDeleteReviewFolder();
    setStatus(
      `Moving ${allIds.length} message${allIds.length === 1 ? "" : "s"} from ${addresses.length} sender${addresses.length === 1 ? "" : "s"}…`
    );

    const { moved, failed } = await moveMessages(
      allIds,
      folderId,
      (done, total) => {
        const pct = Math.round((done / total) * 100);
        els["bar-fill"].style.width = `${pct}%`;
        els["progress-text"].textContent = `${done} / ${total}`;
      }
    );

    if (failed.length === 0) {
      setStatus(
        `Done. Moved ${moved} message${moved === 1 ? "" : "s"} from ${addresses.length} sender${addresses.length === 1 ? "" : "s"} to Delete Review.`,
        "success"
      );
    } else {
      setStatus(
        `Moved ${moved}. ${failed.length} could not be moved — scan again to retry.`,
        "error"
      );
    }

    // Drop the swept senders from the list so it reflects what's left.
    ranked = ranked.filter((s) => !addresses.includes(s.address));
    renderSenderList(ranked);
    els["sender-list"].hidden = ranked.length === 0;
    els.sweep.hidden = ranked.length === 0;
    updateSweepButton();
  } catch (err) {
    setStatus(err.message || String(err), "error");
    els.sweep.disabled = false;
  } finally {
    els.scan.disabled = false;
    els.find.disabled = false;
  }
}

async function onMove() {
  if (matched.length === 0) return;
  els.move.disabled = true;
  els.find.disabled = true;
  els["progress-wrap"].hidden = false;
  setStatus("Preparing Delete Review folder…");

  try {
    const folderId = await ensureDeleteReviewFolder();
    const ids = matched.map((m) => m.id);
    setStatus(`Moving ${ids.length} messages…`);

    const { moved, failed } = await moveMessages(ids, folderId, (done, total) => {
      const pct = Math.round((done / total) * 100);
      els["bar-fill"].style.width = `${pct}%`;
      els["progress-text"].textContent = `${done} / ${total}`;
    });

    if (failed.length === 0) {
      setStatus(
        `Done. Moved ${moved} message${moved === 1 ? "" : "s"} to Delete Review.`,
        "success"
      );
    } else {
      setStatus(
        `Moved ${moved}. ${failed.length} could not be moved — try Find again to retry.`,
        "error"
      );
    }
    resetResults();
    els.sender.value = "";
  } catch (err) {
    setStatus(err.message || String(err), "error");
    els.move.disabled = false;
  } finally {
    els.find.disabled = false;
  }
}

async function boot() {
  if (!CLIENT_ID || CLIENT_ID === "PASTE_YOUR_CLIENT_ID_HERE") {
    showSignedOut();
    setStatus(
      "Not configured yet: set VITE_MS_CLIENT_ID (or edit src/config.js) with your Azure app's client id.",
      "error"
    );
    els.signin.disabled = true;
    return;
  }

  els.signin.addEventListener("click", () => login());
  els.signout.addEventListener("click", () => logout());
  els.find.addEventListener("click", onFind);
  els.move.addEventListener("click", onMove);
  els.scan.addEventListener("click", onScan);
  els.sweep.addEventListener("click", onSweep);
  els.sender.addEventListener("keydown", (e) => {
    if (e.key === "Enter") onFind();
  });

  try {
    const account = await initAuth();
    if (account) {
      showSignedIn(account);
      try {
        const me = await getProfile();
        if (me) els["account-name"].textContent = me.mail || me.userPrincipalName;
      } catch {
        /* profile is cosmetic; ignore */
      }
      setStatus("Connected. Enter a sender to sweep.", "info");
    } else {
      showSignedOut();
    }
  } catch (err) {
    showSignedOut();
    setStatus(err.message || String(err), "error");
  }
}

boot();
