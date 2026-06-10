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
} from "./graph.js";
import { CLIENT_ID } from "./config.js";

const $ = (id) => document.getElementById(id);
const els = {};
[
  "account", "account-name", "signout", "signin-card", "signin", "app-card",
  "sender", "find", "results", "count", "count-sender", "preview", "move",
  "progress-wrap", "bar-fill", "progress-text", "status",
].forEach((id) => (els[id] = $(id)));

let matched = []; // currently found messages awaiting a move

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
