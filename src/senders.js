// Pure aggregation/scoring logic for the sender picker. No network here —
// it just turns a flat list of scanned messages into a ranked list of
// candidate senders, so it's easy to reason about (and to tweak the scoring).

import {
  BULK_ADDRESS_HINTS,
  BULK_SUBDOMAIN_HINTS,
  SPAM_SUBJECT_HINTS,
} from "./config.js";

// messages: [{ id, from, subject, receivedDateTime }]
// Returns: [{ address, name, count, latest, sampleSubject, score, tags }]
// sorted most-spammy first.
export function rankSenders(messages) {
  const map = new Map();

  for (const m of messages) {
    const ea = m.from && m.from.emailAddress;
    if (!ea) continue;
    const address = (ea.address || "").toLowerCase();
    if (!address) continue;

    let s = map.get(address);
    if (!s) {
      s = { address, name: ea.name || "", count: 0, latest: "", subjects: [] };
      map.set(address, s);
    }
    s.count += 1;
    if (!s.name && ea.name) s.name = ea.name;
    if (m.receivedDateTime && m.receivedDateTime > s.latest) {
      s.latest = m.receivedDateTime;
    }
    if (m.subject && s.subjects.length < 3) s.subjects.push(m.subject);
  }

  const list = [...map.values()].map((s) => {
    s.sampleSubject = s.subjects[0] || "(no subject)";
    const { score, tags } = scoreSender(s);
    s.score = score;
    s.tags = tags;
    delete s.subjects;
    return s;
  });

  list.sort((a, b) => b.score - a.score || b.count - a.count);
  return list;
}

// Lean scoring: volume is the dominant signal, with small nudges from
// bulk-style addresses, sending subdomains, and promo-ish subjects.
function scoreSender(s) {
  const tags = [];
  let score = s.count; // volume first

  if (s.count >= 10) tags.push("High volume");

  const [local, domain = ""] = s.address.split("@");

  if (BULK_ADDRESS_HINTS.some((h) => local === h || local.includes(h))) {
    score += 6;
    tags.push("Bulk address");
  }

  if (BULK_SUBDOMAIN_HINTS.some((p) => domain.startsWith(p))) {
    score += 3;
    tags.push("Sending subdomain");
  }

  const subj = (s.sampleSubject || "").toLowerCase();
  if (SPAM_SUBJECT_HINTS.some((k) => subj.includes(k))) {
    score += 1;
    tags.push("Promo subject");
  }

  return { score, tags };
}
