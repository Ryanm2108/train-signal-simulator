// Build-time Excel -> JSON for Route DB (fixed columns 5,6,25; robust multi-row + AGGN handling)
import * as XLSX from "xlsx";
import fs from "fs";
import path from "path";

// ===== config =====
const SRC_XLSX  = process.argv[2] || "data/route-control-table.xlsx";
const OUT_JSON  = process.argv[3] || "src/data/route-db.json";
const SHEET_NAME = "3.Route Control Table"; // change if your sheet name differs

// Fixed columns (Excel 1-based -> 0-based)
const SIGNAL_IDX   = 5 - 1;   // Signal Button
const ROUTE_IDX    = 6 - 1;   // Route Button
const RELEASED_IDX = 25 - 1;  // ROUTE RELEASED TRACK SECTIONS OCCUPIED/CLEARED

// ===== helpers =====

// Allowed route signals: strictly S#, C#, or SH# (prevents "SIGNAL" header from passing)
function startsWithAllowedSignal(value) {
  const s = String(value ?? "").replace(/\s+/g, " ").trim().toUpperCase();
  return /^(?:SH|S|C)\d+/.test(s);
}

// Track-section like 203AXT / 01XT / 202BXT / 203AXT etc.
const TS_RE          = /\b[0-9A-Z]+(?:[AB])?XT\b/gi;
// Prefer “AFTER 120 SEC”; fallback handles “120 SEC” (just in case)
const AFTER_RE       = /after\s+(\d+)\s*sec/i;
const PLAIN_SEC_RE   = /(\d+)\s*sec/i;

// normalize funky whitespace
function clean(s) {
  return String(s ?? "")
    .replace(/\u00A0/g, " ")    // NBSP -> space
    .replace(/\u200B/g, "")     // zero-width space
    .replace(/\s+/g, " ")
    .trim();
}

function isBlankLike(v) {
  if (v === null || v === undefined) return true;
  const s = clean(v);
  return s === "" || s === "-" || s === "–" || s === "—";
}

// Normalize "sig-route" keys to avoid duplicates from stray spaces/casing (e.g., "S1- 02AA")
function makeKey(sig, rte) {
  const s = clean(sig).toUpperCase();
  const r = clean(rte).toUpperCase();
  return `${s}-${r}`.replace(/\s*-\s*/g, "-").replace(/\s+/g, "");
}

// Merge helpers to prevent clobbering non-empty with empty and to dedupe
function haveData(x) {
  return (x?.releaseGroups?.length ?? 0) > 0 || (x?.timedRelease ?? null) !== null;
}

function dedupeGroups(arr) {
  const seen = new Set();
  const norm = g => g.map(x => x.toUpperCase()).sort().join(",");
  const out = [];
  for (const g of arr) {
    const k = norm(g);
    if (!seen.has(k)) { seen.add(k); out.push(g); }
  }
  return out;
}

/**
 * Parse an array of text lines from col 25 (one per *physical* row),
 * stitching groups that span multiple rows:
 *   "(1XT" + "202BXT)" => ["1XT","202BXT"]
 *   "(203BXT)"         => ["203BXT"]
 *
 * Timed case:
 *   e.g. "(203AXT AFTER 120 SEC)" → timedRelease: { track:"203AXT", durationSec:120 }
 *   We only set timedRelease if that *same* line has "SEC" and exactly one TS id.
 */
function parseReleaseLines(lines) {
  const groups = [];
  let timedRelease = null;

  // tolerant: "(203AXT AFTER 120 SEC)", "(01XT [TS] 120 SEC.)", "(01XT 90 SECS)"
  const TIMED_RE =
    /\(\s*([0-9A-Z]+(?:[AB])?XT)\s*(?:\[\s*TS\s*\])?(?:\s*AFTER)?\s*(\d+)\s*SECS?\.?\s*\)/i;

  const pushGroup = (ids) => {
    if (!ids || !ids.length) return;
    const key = ids.map(x => x.toUpperCase()).sort().join(",");
    if (!pushGroup._seen) pushGroup._seen = new Set();
    if (!pushGroup._seen.has(key)) {
      pushGroup._seen.add(key);
      groups.push(ids);
    }
  };

  // Stitch text between '(' and ')' that may span multiple rows/cells
  let parenOpen = false;
  let parenBuf = "";

  const finalizeParen = () => {
    if (!parenBuf) return;

    // 1) groups from the full paren chunk
    const ids = (parenBuf.match(TS_RE) || []).map(x => x.toUpperCase());
    if (ids.length) pushGroup(ids);

    // 2) timed seconds from the full paren chunk (keep first)
    if (timedRelease === null) {
      const m = parenBuf.match(TIMED_RE);
      if (m) {
        timedRelease = parseInt(m[2], 10); // <-- NUMBER ONLY
      }
    }

    parenOpen = false;
    parenBuf = "";
  };

  for (const raw of lines) {
    if (isBlankLike(raw)) continue;

    const pieces = String(raw)
      .split(/\r?\n|\r|\u000b|\u2028|\u2029/g)
      .map(clean)
      .filter(Boolean);

    for (const line of pieces) {
      const starts = /^\s*\(/.test(line);
      const ends   = /\)\s*$/.test(line);

      if (starts && ends) {
        parenBuf = line;
        finalizeParen();
        continue;
      }
      if (starts && !ends) {
        parenOpen = true;
        parenBuf = line;
        continue;
      }
      if (!starts && ends && parenOpen) {
        parenBuf += " " + line;
        finalizeParen();
        continue;
      }
      if (parenOpen) {
        parenBuf += " " + line;
        continue;
      }

      // No parentheses here: treat any TS as singleton groups
      const ids = (line.match(TS_RE) || []).map(x => x.toUpperCase());
      ids.forEach(id => pushGroup([id]));
    }
  }

  if (parenOpen) finalizeParen(); // close dangling chunk (defensive)

  return { groups, timedRelease }; // number or null
}



// ===== read workbook =====
const buf = fs.readFileSync(SRC_XLSX);
const wb  = XLSX.read(buf, { type: "buffer" });

let ws = SHEET_NAME ? wb.Sheets[SHEET_NAME] : null;
if (!ws) {
  const guess = wb.SheetNames.find(n => n.toUpperCase().includes("ROUTE CONTROL TABLE"));
  ws = wb.Sheets[guess || wb.SheetNames[0]];
}
if (!ws) throw new Error("Route sheet not found.");

// defval: "" ensures we get cells instead of undefined, which keeps indexes stable
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, raw: false, defval: "" });
if (!rows.length) throw new Error("Empty sheet.");

// ===== find first data row automatically (avoid header/labels) =====
let startRowIdx = 0;
for (let r = 0; r < rows.length; r++) {
  const sig = clean(rows[r]?.[SIGNAL_IDX]);
  const rte = clean(rows[r]?.[ROUTE_IDX]);
  const looksHeaderish = /BUTTON|SIGNAL|ROUTE/i.test(`${sig} ${rte}`);
  if (startsWithAllowedSignal(sig) && rte && !looksHeaderish) {
    startRowIdx = r;
    break;
  }
}
console.log(`ℹ️ Using cols (0-based): signal=${SIGNAL_IDX}, route=${ROUTE_IDX}, released=${RELEASED_IDX}`);
console.log(`ℹ️ Data starts at rowIndex=${startRowIdx} (Excel row ${startRowIdx + 1})`);

// ===== build DB by blocks =====
const out = {};
for (let r = startRowIdx; r < rows.length; ) {
  const row = rows[r] || [];
  const sig = (row[SIGNAL_IDX] ?? "").toString().trim();
  const rte = (row[ROUTE_IDX] ?? "").toString().trim();

  // ⛔ stop the entire read when Signal doesn't start with S/C/SH (e.g., EP/DP blocks)
  if (sig && !startsWithAllowedSignal(sig)) {
    console.log(`🛑 Reached non-route section at row ${r + 1} (signal="${sig}") — stopping.`);
    break;
  }

  if (!sig || !rte) { r++; continue; }          // skip non-anchor rows

  if (rte.toUpperCase().includes("AGGN")) {
    console.log(`⏭️  Skipping AGGN variant anchor at row ${r+1}: ${sig}-${rte}`);
    r++;
    continue;
  }

  const key = makeKey(sig, rte);

  // collect all lines for col 25 for this block,
  // including AGGN row(s) as continuation if they appear directly after
  const lines = [];
  let j = r;
  while (j < rows.length) {
    const rr   = rows[j] || [];
    const sigJ = (rr[SIGNAL_IDX] ?? "").toString().trim();
    const rteJ = (rr[ROUTE_IDX]  ?? "").toString().trim();

    // ⛔ non-route section (e.g., EP/DP) is a hard boundary even if not an anchor
    if (sigJ && !startsWithAllowedSignal(sigJ)) break;

    // ➜ The next *non-AGGN* anchor row starts a new block
    const isAnchor = !!(sigJ && rteJ);
    const isAggn   = rteJ.toUpperCase().includes("AGGN");
    if (j > r && isAnchor && !isAggn) break;

    // We do NOT break on AGGN; we still harvest its col-25 lines if they spill from the normal row.
    const rel = rr[RELEASED_IDX];
    if (!isBlankLike(rel)) lines.push(rel);

    j++;
  }

  const { groups, timedRelease } = parseReleaseLines(lines);
  const parsed = { releaseGroups: groups, timedRelease: timedRelease ?? null };

  if (out[key]) {
    // duplicate / near-duplicate (e.g., spacing variants) — merge safely
    if (!haveData(parsed) && haveData(out[key])) {
      console.warn(`⚠️ Duplicate ${key} at row ${r+1}: incoming empty → keeping existing.`);
    } else {
      out[key] = {
        releaseGroups: dedupeGroups([...(out[key].releaseGroups || []), ...(parsed.releaseGroups || [])]),
        timedRelease: out[key].timedRelease ?? parsed.timedRelease ?? null,
      };
      if (
        out[key].timedRelease &&
        parsed.timedRelease &&
        (out[key].timedRelease.track !== parsed.timedRelease.track ||
         out[key].timedRelease.durationSec !== parsed.timedRelease.durationSec)
      ) {
        console.warn(`⚠️ Conflicting timedRelease for ${key}; keeping first.`);
      }
      console.warn(`⚠️ Merged duplicate ${key} at row ${r+1}.`);
    }
  } else {
    out[key] = parsed;
  }

  r = j; // jump to next route block
}

// ===== write JSON =====
fs.mkdirSync(path.dirname(OUT_JSON), { recursive: true });
fs.writeFileSync(OUT_JSON, JSON.stringify(out, null, 2));
console.log(`✅ Wrote ${Object.keys(out).length} routes → ${OUT_JSON}`);
