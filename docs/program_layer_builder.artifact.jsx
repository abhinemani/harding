import { useState, useMemo, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

// ---- palette (Public Works design system) -------------------------------
const C = {
  maroon: "#6B1A2A", maroonDark: "#4E121E", gold: "#C99C1C", sand: "#F4EEE2",
  ink: "#2C2C2C", grey: "#6B6B6B", line: "#D9D0BF", paper: "#FBF8F1", white: "#FFFFFF",
};
const serif = "Georgia, 'Times New Roman', serif";
const sans = "Arial, Helvetica, sans-serif";

// ---- helpers -------------------------------------------------------------
function parseAmount(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return v;
  let s = String(v).trim();
  if (!s) return null;
  const neg = /^\(.*\)$/.test(s) || s.startsWith("-");
  s = s.replace(/[()$,\s-]/g, "");
  if (!s) return 0; // "$ -" means zero in the county export
  if (isNaN(Number(s))) return null;
  return neg ? -Number(s) : Number(s);
}
const norm = (h) => String(h ?? "").replace(/\s+/g, " ").trim();
const money = (n) =>
  n === null || n === undefined || isNaN(n)
    ? "—"
    : new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(n);

// Column defaults follow the county's standard export; pattern-match only as a fallback.
function guessColumns(headers) {
  const lower = headers.map((h) => h.toLowerCase());
  const exact = (n) => { const i = lower.indexOf(n.toLowerCase()); return i >= 0 ? headers[i] : ""; };
  const find = (tests, exclude = /$^/) => {
    for (const t of tests) { const i = lower.findIndex((h) => t.test(h) && !exclude.test(h)); if (i >= 0) return headers[i]; }
    return "";
  };
  return {
    cc: exact("Costing Center Name") || find([/costing ?center/, /cost ?center/, /program/], /acct|code|ref/),
    gl: exact("GL Account Name") || find([/account ?name/, /object/, /desc/, /line ?item/]),
    code: exact("GL Account Number") || find([/account ?(#|no|num)/, /^gl/, /object ?code/]),
    amt: exact("Adopted Budget") || find([/adopted/, /approved/, /proposed/, /budget/, /amount/]),
    div: exact("Division Name") || find([/division/, /agency/, /department/], /code/),
    fund: exact("Fund Code") || find([/^fund/]),
  };
}

// ---- the prompt ----------------------------------------------------------
// Encodes the Program Standard (Public Works LLC / Funkhouser & Associates, Line of Sight) and what the
// four pilot divisions' self-inventories taught us about how divisions misdescribe their own work.
function buildPrompt(division, items, total, personnelShare, emptyCenters, headcount, recoveries) {
  const lines = items.map((it) => `${it.id}\t${it.fund || ""}\t${it.cc}\t${it.code}\t${it.desc}\t${Math.round(it.amt)}`).join("\n");
  return `You are a public-finance analyst helping Frederick County, Maryland (pop. ~290,000, ~$1.55B operating budget) add a program layer on top of its line-item budget. The general ledger is untouched; a program is an additive classification. Your job is to draft this division's programs to the County's Program Standard, reproduced below. Follow it strictly on definition, boundaries, wording, and measures.

Division: ${division || "(unnamed)"}. Infer what it does from its name and the costing center and account names, using your knowledge of how Maryland county governments are organized. Decide whether it is external-facing (residents are the beneficiaries) or internal (other divisions or employees are), and decide the same for each program, since some divisions hold both.

The County's model splits evidence by orientation, because a single standard fails everyone:
- INTERNAL programs are budgeted on PERFORMANCE — service measures the division itself controls: timeliness, reliability, backlog, cost per transaction (HR's time-to-fill, cost per hire; IT's incidents resolved within target). Do not borrow public outcomes for them.
- EXTERNAL programs are budgeted on OUTCOMES — the community result the program exists to advance (emissions reduced, older adults remaining at home, pollutant loads declining). State the program's intended contribution; no one pretends a single program controls a community-wide condition.

Line items — id, fund code, costing center, GL account number, GL account name, FY adopted budget (revenue, zero, and 6xxxx cost-recovery lines already removed). Fund 10 is the General Fund; 2x funds are grants; other funds are usually enterprise or internal service funds that bill customers for a discrete service and should generally be their own program. Total: ${Math.round(total)}. Personnel accounts (50xxx) are ${Math.round(personnelShare * 100)}% of the total; at Frederick County's average loaded cost that implies roughly ${headcount} staff.
${lines}
${recoveries.length ? `\nCost recoveries (6xxxx) set aside from the total — money this division bills to other units for services it provides, which tells you who its customers are: ${recoveries.map((r) => `${r.desc} ${Math.round(r.amt)}`).join("; ")}. When a recovery equals a costing center's spending, that program is fee-for-service; its net general-fund cost is near zero and its 20% answer should speak to capacity and rate adequacy.` : ""}
${emptyCenters.length ? `\nCosting centers in the file carrying $0 this year (usually grants active in prior years or awaiting award — evidence of what the division does, not of what this budget funds): ${emptyCenters.join("; ")}` : ""}

THE PROGRAM STANDARD
Definition. A program is a set of activities that delivers a defined result to a defined beneficiary, carries its full cost within one division, and could be funded at a different level next year with an observable consequence. If you cannot name who is better off and how, you have named a cost center. If funding it at 80% would change nothing describable, the boundary is wrong.

Boundary rules.
1. Exclusive and exhaustive: every dollar in this file sits in exactly one program and the programs sum to the total. No zero-cost rows.
2. Cost follows the budget; accountability is noted separately. Work performed or paid for elsewhere (another division, a countywide non-departmental account, CIP, bonds, grants not in this file, a vendor or carrier) is NOT a program of this division. Record such accountabilities in accountabilities_elsewhere with where the money sits, and note contributing divisions on the programs they touch. This overrides the instinct to list benefit plans, permit obligations, or capital projects the division only administers.
3. Span: no program below ~3% or above ~25% of the division total. Most divisions land between 5 and 15 programs. A row holding a third of the budget and most of the staff is a costing center wearing a program's name; split it into the services it delivers.
4. Administration is named, not smuggled: at most one administration program, near or under 10%; allocate anything above that to the programs it supports.

What divisions get wrong when they self-inventory (all observed in Frederick): they list the visible things (benefit plans, named initiatives, permit chapters, technology domains, policy goals) and omit the core operating work that consumes most of the headcount — represent the core work first. They carry one figure across several rows. They list obligations funded elsewhere at $0. They put a third of the budget in one undifferentiated row. They name programs for statutes and funding streams. They offer deadlines, activity counts, and "remain in compliance" as measures.

Program statement. One sentence to a fixed grammar: "[Name] delivers [service] to [beneficiary] so that [result], at [$ cost] and [FTE] in FY26." Name the program for the service, never the office, statute, or funding stream ("Home-Delivered and Congregate Nutrition", not "Title III-C"). Avoid effort verbs — support, manage, oversee, coordinate, administer, facilitate, maintain — except when the object is a physical asset; use verbs of delivery: provides, serves, replaces, treats, houses, trains, inspects, restores. The result clause states the condition that changes. "So that the County remains in compliance" is a constraint, not a result; where compliance is the purpose, state the condition it protects. Omit the "at $ and FTE" tail — the app computes and appends it.

Measures. Exactly three per program, of three kinds: a RESULT measure — for an external program, the community outcome it contributes to; for an internal program, the service level it controls (timeliness, reliability, quality) — a VOLUME measure (how much was delivered), and a UNIT COST (cost ÷ volume, using the volume measure). Each must come from a system that already produces it (Munis, HRIS, applicant tracking, CRM, work orders, permit and inspection systems, carrier or actuary reports, state reports); if the number does not exist today, name the closest proxy and rate exists_today honestly. A deadline is not a measure.

Mandate status: one of "required by law" (cite it), "required to receive funding" (name the source), or "discretionary". Priority alignment: County leadership will set 5–7 priorities anchored in Livable Frederick; until then, name the specific Livable Frederick element, adopted plan, or Executive priority the program advances, and rate alignment as "direct" (the program exists to move it), "contributes" (it helps, among other things), or "none" (an essential or mandated service that aligns to no priority — a legitimate and important answer, since the point is to see which resources could realistically be redirected and which cannot). An honest "none" beats a stretch. The twenty percent question: what specifically stops at −20%, and what becomes possible at +20% — two concrete sentences, in specifics not generalities.

COSTING
- Assign a line WHOLLY to one program via item_ids, or, if genuinely shared (salaries, benefits, telephone, supplies in one costing center), give each program its share of that line in its own "shared" list. A line's shares across all programs should sum to 1. Never put the same line in item_ids and shared.
- Emit the programs in descending order of cost, largest first, so the reader sees the division's shape early.
- When most of the budget is personnel in one costing center, the program layer is a staff-time allocation. Propose shares from what a division of this kind spends its people on, sized against the implied headcount; benefits follow salary shares. State the logic in allocation_note as a starting bid for the director to correct. Give fte per program consistent with the shares and the implied headcount, to two decimals.
- When the file has many costing centers named for grants, titles, or funded services, each is already a program component; roll them up into service programs and allocate the unnamed general-fund center across them. Say which costing centers each program absorbs.
- If one non-personnel line exceeds a quarter of the budget (a software or subscription pool, a large contract), it is a portfolio bought for several programs. Allocate it provisionally and say in allocation_note that the real split needs the underlying vendor or subscription list; put that in data_requests.
- If the division is organized around a permit, mandate, or statutory frame, say so in division_read and note in each program which elements of the frame it covers; the program layer is coarser than the frame and crosswalks to it.

Keep strings terse: statement under 35 words, measures under 10, sources under 5, notes under 30, each 20% sentence under 25 words. Omit null fields. Return ONLY a JSON object, no fences, no prose:
{
  "division_read": "one sentence, under 35 words: what the line items reveal about how this division actually spends money",
  "orientation": "external" | "internal" | "mixed",
  "programs": [
    {
      "name": "service-named program",
      "facing": "external" | "internal",
      "statement": "[Name] delivers [service] to [beneficiary] so that [result]",
      "beneficiary": "who, roughly how many, where",
      "fte": 4.25,
      "item_ids": [3, 9],
      "shared": [{"item_id": 7, "share": 0.4}],
      "allocation_note": "",
      "contributing_divisions": "",
      "measures": {
        "result": {"metric": "...", "source": "...", "exists_today": "likely"|"possible"|"unlikely"},
        "volume": {"metric": "...", "source": "...", "exists_today": "likely"|"possible"|"unlikely"},
        "unit_cost": {"metric": "cost per ...", "source": "...", "exists_today": "likely"|"possible"|"unlikely"}
      },
      "mandate": {"status": "required by law"|"required to receive funding"|"discretionary", "citation": "statute, funding source, or empty"},
      "strategic_linkage": "specific element or 'none identified'",
      "alignment": "direct" | "contributes" | "none",
      "twenty_percent": {"minus": "what stops at −20%", "plus": "what becomes possible at +20%"},
      "is_admin": false
    }
  ],
  "accountabilities_elsewhere": [{"name": "...", "where": "CIP | grants | non-departmental | another division | vendor/carrier", "note": "one line"}],
  "unassigned_ids": [],
  "cautions": ["at most 3 short notes on where this is a judgment call or the data is thin"],
  "data_requests": ["at most 3 specific things to ask the division for that would turn provisional allocations into real ones"]
}
Emit "division_read" and "orientation" first, then "programs".`;
}

// Pull whatever is complete out of a partially streamed JSON reply: the one-line read, orientation, and every
// program object whose closing brace has arrived. Respects strings so braces inside text don't fool it.
function parsePartial(text) {
  const out = { division_read: "", orientation: "", programs: [] };
  const dr = text.match(/"division_read"\s*:\s*"((?:[^"\\]|\\.)*)"/); if (dr) { try { out.division_read = JSON.parse(`"${dr[1]}"`); } catch { out.division_read = dr[1]; } }
  const or = text.match(/"orientation"\s*:\s*"(external|internal|mixed)"/); if (or) out.orientation = or[1];
  const start = text.indexOf('"programs"'); if (start < 0) return out;
  let i = text.indexOf("[", start); if (i < 0) return out;
  i += 1;
  while (i < text.length) {
    while (i < text.length && /[\s,]/.test(text[i])) i++;
    if (text[i] !== "{") break;
    let depth = 0, inStr = false, esc = false, j = i;
    for (; j < text.length; j++) {
      const ch = text[j];
      if (inStr) { if (esc) esc = false; else if (ch === "\\") esc = true; else if (ch === '"') inStr = false; continue; }
      if (ch === '"') inStr = true; else if (ch === "{") depth++; else if (ch === "}") { depth--; if (depth === 0) break; }
    }
    if (depth !== 0) break; // this program is still arriving
    try { out.programs.push(JSON.parse(text.slice(i, j + 1))); } catch { break; }
    i = j + 1;
  }
  return out;
}

// ---- component -----------------------------------------------------------
export default function ProgramLayerBuilder() {
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [cols, setCols] = useState({ cc: "", gl: "", code: "", amt: "", div: "", fund: "" });
  const [division, setDivision] = useState("");
  const [status, setStatus] = useState("idle");
  const [progress, setProgress] = useState("");
  const [draft, setDraft] = useState(null);
  const [stage, setStage] = useState("idle"); // idle | reading | thinking | drafting | finishing | done
  const [setupOpen, setSetupOpen] = useState(true);
  const [pools, setPools] = useState({});
  const [sendCount, setSendCount] = useState(0);
  const [mdText, setMdText] = useState("");
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);
  const [open, setOpen] = useState({});
  const [narrow, setNarrow] = useState(() => typeof window !== "undefined" && window.innerWidth < 760);
  useEffect(() => {
    const mq = window.matchMedia("(max-width: 760px)");
    const fn = (e) => setNarrow(e.matches); setNarrow(mq.matches);
    mq.addEventListener ? mq.addEventListener("change", fn) : mq.addListener(fn);
    return () => (mq.removeEventListener ? mq.removeEventListener("change", fn) : mq.removeListener(fn));
  }, []);
  const pad = narrow ? 16 : 22;
  const [csvText, setCsvText] = useState("");
  const [copied, setCopied] = useState(false);

  async function onBudget(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    setFileName(f.name); setResult(null); setError(""); setStatus("parsing");
    try {
      let data = [];
      if (/\.csv$/i.test(f.name)) {
        data = Papa.parse(await f.text(), { header: true, skipEmptyLines: true, transformHeader: norm }).data;
      } else {
        const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
        data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        data = data.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [norm(k), v])));
      }
      const hdrs = Object.keys(data[0] || {});
      const g = guessColumns(hdrs);
      setHeaders(hdrs); setRows(data); setCols(g);
      if (g.div && !division) {
        const d = data.map((r) => norm(r[g.div])).find(Boolean) || "";
        setDivision(d.replace(/^division of\s+/i, ""));
      }
      setStatus("idle");
    } catch { setError("Couldn't read the budget file. The first row must be column headers."); setStatus("error"); }
  }

  const items = useMemo(() => {
    if (!rows.length || !cols.amt || (!cols.gl && !cols.cc)) return [];
    const out = [];
    rows.forEach((r, i) => {
      const amt = parseAmount(r[cols.amt]);
      const code = cols.code ? norm(r[cols.code]) : "";
      const cc = cols.cc ? norm(r[cols.cc]) : "";
      const gl = cols.gl ? norm(r[cols.gl]) : "";
      if (amt === null || amt === 0) return;
      if (code.startsWith("4") || code.startsWith("6")) return; // revenue, and 6xxxx cost recoveries
      const fund = cols.fund ? norm(r[cols.fund]) : "";
      out.push({ id: out.length + 1, row: i, cc, code, fund, desc: gl || cc, amt, personnel: code.startsWith("50") });
    });
    return out;
  }, [rows, cols]);
  const recoveries = useMemo(() => {
    if (!rows.length || !cols.code) return [];
    return rows.map((r) => ({ code: norm(r[cols.code]), desc: cols.gl ? norm(r[cols.gl]) : "", cc: cols.cc ? norm(r[cols.cc]) : "", amt: parseAmount(r[cols.amt]) }))
      .filter((x) => x.code.startsWith("6") && x.amt);
  }, [rows, cols]);
  const recoveryTotal = recoveries.reduce((s, x) => s + x.amt, 0);

  const total = useMemo(() => items.reduce((s, it) => s + it.amt, 0), [items]);
  const personnelShare = useMemo(() => (total ? items.filter((i) => i.personnel).reduce((s, i) => s + i.amt, 0) / total : 0), [items, total]);
  const zeroCount = rows.length - items.length;
  const baseName = (cc) => cc.replace(/\b(FY ?\d{2,4}|20\d{2}|Rd ?#?\d+|-?New#|- ?Copy)\b/gi, "").replace(/[\s\-#]+$/g, "").replace(/\s{2,}/g, " ").trim();
  const emptyCenters = useMemo(() => {
    if (!rows.length || !cols.cc) return [];
    const funded = new Set(items.map((i) => baseName(i.cc)));
    const seen = new Set();
    rows.forEach((r) => { const b = baseName(norm(r[cols.cc])); if (b && !funded.has(b)) seen.add(b); });
    return [...seen].slice(0, 40);
  }, [rows, cols, items]);
  const headcount = useMemo(() => Math.max(1, Math.round(items.filter((i) => i.personnel).reduce((s, i) => s + i.amt, 0) / 110000)), [items]);

  async function analyze() {
    if (!items.length) return;
    setStatus("thinking"); setError(""); setResult(null); setCsvText(""); setProgress(""); setDraft(null); setStage("reading");
    // Compress within each costing center when the file is large: personnel pooled, small operating lines pooled, big lines kept.
    // Pools carry negative ids; the reply's references to them are expanded back to the real lines below.
    let send = []; const pools = {};
    setPools(pools);
    {
      let pid = -1;
      const byCc = {}; items.forEach((it) => { (byCc[it.cc] ||= []).push(it); });
      Object.entries(byCc).forEach(([cc, its]) => {
        const pers = its.filter((i) => i.personnel), ops = its.filter((i) => !i.personnel);
        const ccTotal = its.reduce((s, i) => s + i.amt, 0);
        if (pers.length) { pools[pid] = pers.map((i) => i.id); send.push({ id: pid--, cc, fund: pers[0].fund, code: "50xxx", desc: `Personnel — salary & benefits (${pers.length} lines)`, amt: pers.reduce((s, i) => s + i.amt, 0) }); }
        const big = ops.filter((i) => i.amt >= Math.max(5000, ccTotal * 0.02, total * 0.003)), small = ops.filter((i) => !big.includes(i));
        big.forEach((i) => send.push(i));
        if (small.length) {
          pools[pid] = small.map((i) => i.id);
          const inside = [...small].sort((a, b) => b.amt - a.amt).slice(0, 8).map((i) => `${i.desc} ${Math.round(i.amt / 1000)}k`).join(", ");
          send.push({ id: pid--, cc, fund: small[0].fund, code: "5xxxx", desc: `Other operating, ${small.length} lines: ${inside}${small.length > 8 ? ", …" : ""}`, amt: small.reduce((s, i) => s + i.amt, 0) });
        }
      });
    }
    setSendCount(send.length);
    const expand = (id) => (pools[id] ? pools[id] : [id]);
    const run = async (extra) => {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 12000, stream: true,
          messages: [{ role: "user", content: buildPrompt(division, send, total, personnelShare, emptyCenters, headcount, recoveries) + extra }] }),
      });
      if (!res.ok) { const err = await res.json().catch(() => ({})); throw new Error(`API ${res.status}: ${err.error?.message || res.statusText}`); }
      let text = "", stopReason = null, lastDraft = 0;
      const reader = res.body.getReader(); const dec = new TextDecoder(); let buf = "";
      while (true) {
        const { value, done } = await reader.read(); if (done) break;
        buf += dec.decode(value, { stream: true });
        const parts = buf.split("\n"); buf = parts.pop();
        for (const line of parts) {
          if (!line.startsWith("data:")) continue;
          let ev; try { ev = JSON.parse(line.slice(5)); } catch { continue; }
          if (ev.type === "content_block_delta" && ev.delta?.text) {
            text += ev.delta.text;
            const now = Date.now();
            if (now - lastDraft > 200) {
              lastDraft = now;
              const d = parsePartial(text);
              setDraft(d);
              setStage(text.includes('"accountabilities_elsewhere"') || text.includes('"cautions"') ? "finishing" : "drafting");
              const n = d.programs.length;
              const drafting = (text.match(/"name":/g) || []).length;
              setProgress(n ? `${n} program${n > 1 ? "s" : ""} drafted${drafting > n ? ", working on the next" : ""}…` : drafting ? "Drafting the first program…" : "Reading the budget…");
            }
          }
          if (ev.type === "message_delta" && ev.delta?.stop_reason) stopReason = ev.delta.stop_reason;
          if (ev.type === "error") throw new Error(ev.error?.message || "stream error");
        }
      }
      return { text, stopReason };
    };
    try {
      let { text, stopReason } = await run("");
      if (stopReason === "max_tokens") {
        setProgress("The first draft ran long; asking for a tighter one…"); setDraft(null);
        ({ text, stopReason } = await run("\n\nIMPORTANT: your previous reply exceeded the length limit and was discarded. Produce the same JSON at half the length: at most 6 programs, 2 outputs and 1 outcome each, every string under 15 words, notes under 20 words, no more than 2 cautions and 2 data_requests."));
      }
      if (stopReason === "max_tokens") throw new Error("The reply was cut off before the JSON finished, even after a retry. Try a smaller file or a different model.");
      const s = text.indexOf("{"), e = text.lastIndexOf("}");
      if (s < 0) throw new Error(`No JSON in reply. It began: ${text.slice(0, 160)}`);
      let parsed;
      try { parsed = JSON.parse(text.slice(s, e + 1)); } catch (err) { throw new Error(`Couldn't parse the reply (${err.message}). It began: ${text.slice(0, 160)}`); }
      setResult(parsed); setDraft(null); setStatus("done"); setStage("done"); if (narrow) setSetupOpen(false);
    } catch (err) { setError(err?.message || "The analysis didn't come back as expected."); setStatus("error"); setDraft(null); setStage("idle"); }
  }

  // ---- costing from whichever we have: the finished reply, or the programs streamed so far
  const live = result || draft;
  const streaming = !result && !!draft;
  const byId = useMemo(() => Object.fromEntries(items.map((it) => [it.id, it])), [items]);
  const programs = useMemo(() => {
    if (!live) return [];
    const expand = (id) => (pools[id] ? pools[id] : [id]);
    const progs = live.programs.map((p) => ({
      ...p,
      item_ids: (p.item_ids || []).flatMap(expand),
      shared: (p.shared || []).flatMap((sh) => expand(sh.item_id).map((id) => ({ item_id: id, share: Math.max(0, Number(sh.share) || 0) }))),
    }));
    // normalize each shared line's shares across programs once the reply is complete; leave raw while streaming
    const sums = {};
    progs.forEach((p) => p.shared.forEach((sh) => { sums[sh.item_id] = (sums[sh.item_id] || 0) + sh.share; }));
    const whole = new Set(progs.flatMap((p) => p.item_ids));
    return progs.map((p) => {
      const direct = p.item_ids.map((id) => byId[id]).filter(Boolean).map((it) => ({ ...it, share: 1, alloc: it.amt }));
      const sharedLines = p.shared.filter((sh) => byId[sh.item_id] && !whole.has(sh.item_id) && sh.share > 0).map((sh) => {
        const norm = streaming ? Math.min(sh.share, 1) : sh.share / (sums[sh.item_id] || 1);
        return { ...byId[sh.item_id], share: norm, alloc: byId[sh.item_id].amt * norm };
      });
      const lines = [...direct, ...sharedLines].sort((a, b) => b.alloc - a.alloc);
      const cost = lines.reduce((s, l) => s + l.alloc, 0);
      return { ...p, lines, cost, share: total ? cost / total : 0 };
    });
  }, [live, byId, total, pools, streaming]);
  const assigned = useMemo(() => new Set(programs.flatMap((p) => p.lines.map((l) => l.id))), [programs]);
  const unassigned = items.filter((it) => !assigned.has(it.id));
  const coverage = total ? (total - unassigned.reduce((s, i) => s + i.amt, 0)) / total : 0;
  const sorted = [...programs].sort((a, b) => b.cost - a.cost);
  const EFFORT_VERBS = /\b(supports?|manages?|oversees?|coordinates?|administers?|facilitates?|maintains?)\b/i;
  const checks = useMemo(() => {
    if (!result) return [];
    const fteSum = programs.reduce((s, p) => s + (Number(p.fte) || 0), 0);
    const admin = programs.filter((p) => p.is_admin);
    const outSpan = programs.filter((p) => p.share < 0.03 || p.share > 0.25);
    const effort = programs.filter((p) => EFFORT_VERBS.test(p.statement || ""));
    const noMeasure = programs.filter((p) => !(p.measures?.result?.metric && p.measures?.volume?.metric && p.measures?.unit_cost?.metric));
    const newData = programs.filter((p) => ["result", "volume", "unit_cost"].some((k) => p.measures?.[k]?.exists_today === "unlikely"));
    const noMandate = programs.filter((p) => !p.mandate?.status);
    const weak20 = programs.filter((p) => !(p.twenty_percent?.minus && p.twenty_percent?.plus));
    const list = (arr) => arr.map((p) => p.name).join(", ");
    return [
      { ok: coverage > 0.995, text: coverage > 0.995 ? "Programs sum to the appropriation." : `${Math.round((1 - coverage) * 100)}% of the budget is unassigned.` },
      { ok: Math.abs(fteSum - headcount) / headcount < 0.25, text: `FTE proposed: ${fteSum.toFixed(2)} against roughly ${headcount} implied by the salary lines.` },
      { ok: outSpan.length === 0, text: outSpan.length ? `Outside the 3–25% span: ${list(outSpan)}.` : "All programs within the 3–25% span." },
      { ok: admin.length <= 1 && admin.every((p) => p.share <= 0.12), text: admin.length ? `Administration: ${list(admin)} at ${admin.map((p) => Math.round(p.share * 100) + "%").join(", ")}.` : "No separate administration program." },
      { ok: effort.length === 0, text: effort.length ? `Effort verbs in the statement: ${list(effort)}.` : "Statements use verbs of delivery." },
      { ok: noMeasure.length === 0, text: noMeasure.length ? `Missing one of result / volume / unit cost: ${list(noMeasure)}.` : "Every program carries a result, a volume, and a unit cost." },
      { ok: newData.length === 0, text: newData.length ? `Measures needing new data collection: ${list(newData)}.` : "Every measure is sourced from a system that exists today." },
      { ok: noMandate.length === 0, text: noMandate.length ? `Mandate status missing: ${list(noMandate)}.` : "Mandate status declared for every program." },
      { ok: weak20.length === 0, text: weak20.length ? `Twenty-percent question unanswered: ${list(weak20)}.` : "Twenty-percent question answered for every program." },
    ];
  }, [result, programs, coverage, headcount]);

  async function exportCsv() {
    const out = [];
    programs.forEach((p) => p.lines.forEach((l) => out.push({
      costing_center: l.cc, gl_account: l.code, gl_account_name: l.desc, adopted_budget: l.amt,
      program: p.name, share: Math.round(l.share * 1000) / 1000, allocated_amount: Math.round(l.alloc),
    })));
    unassigned.forEach((l) => out.push({ costing_center: l.cc, gl_account: l.code, gl_account_name: l.desc, adopted_budget: l.amt, program: "UNASSIGNED", share: 1, allocated_amount: l.amt }));
    const csv = Papa.unparse(out);
    setCsvText(csv);
    try { await navigator.clipboard.writeText(csv); setCopied(true); setTimeout(() => setCopied(false), 2500); } catch { setCopied(false); }
  }

  const [htmlText, setHtmlText] = useState("");
  const esc = (t) => String(t ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  function buildHtml() {
    const today = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
    const dot = (v) => `<span class="dot ${esc(v)}"></span>`;
    const metric = (m) => `<div class="m">${dot(m.exists_today)}<div><div>${esc(m.metric)}</div><div class="src">${esc(m.source)}</div></div></div>`;
    const cards = sorted.map((p, i) => {
      const comp = composition(p.lines).filter((c) => c.amt > 0); const m = p.measures || {};
      const strip = `<div class="strip">${comp.map((c) => `<span style="width:${(c.amt / p.cost) * 100}%;background:${c.color}"></span>`).join("")}</div>
        <div class="comp">${comp.map((c) => `<span><i style="background:${c.color}"></i>${esc(c.label)} ${Math.round((c.amt / p.cost) * 100)}%</span>`).join("")}</div>`;
      const cell = (k, sub, x) => `<div><div class="mh"><b>${k}</b> ${sub}</div>${x?.metric ? `<div class="m">${dot(x.exists_today)}<div><div>${esc(x.metric)}</div><div class="src">${esc(x.source)}</div></div></div>` : `<div class="src" style="color:${C.maroon}">Not proposed.</div>`}</div>`;
      const lines = p.lines.length ? `<div class="build"><div class="lbl">How ${money(p.cost)} is built</div>${p.allocation_note ? `<p class="note">${esc(p.allocation_note)}</p>` : ""}
        <table>${p.lines.map((l) => `<tr><td class="g">${esc(l.code)}</td><td>${esc(l.desc)}${l.cc && l.cc !== l.desc ? ` <span class="g">· ${esc(l.cc)}</span>` : ""}</td><td class="g r">${l.share < 1 ? `${Math.round(l.share * 100)}% of ${money(l.amt)}` : ""}</td><td class="r">${money(l.alloc)}</td></tr>`).join("")}</table></div>` : "";
      const none = p.alignment === "none" || /none identified/i.test(p.strategic_linkage || "");
      const rk = resultKind(p.facing || result.orientation);
      const direct = p.lines.filter((l) => l.share >= 1).reduce((a, l) => a + l.alloc, 0), alloc = p.cost - direct;
      const pack = `<div class="pack"><span class="g">Evidence pack</span>${packItems(p).map(([k, ok]) => `<span style="color:${ok ? C.ink : C.maroon}"><b style="color:${ok ? "#3D6B3A" : C.maroon}">${ok ? "●" : "○"}</b> ${k}</span>`).join("")}</div>`;
      return `<section class="card${p.is_admin ? " admin" : ""}">
        <div class="head"><div><div class="lbl">Program ${i + 1} of ${sorted.length}${p.is_admin ? " · administration" : ""}</div><h2>${esc(p.name)}</h2>
          <p class="stmt">${esc(p.statement)}, at ${money(p.cost)} and ${Number(p.fte || 0).toFixed(2)} FTE in FY26.</p>
          <div class="chips">${p.facing ? `<span class="chip ${p.facing === "internal" ? "teal" : "maroon"}">${esc(facingLabel(p.facing))}</span>` : ""}${p.mandate?.status ? `<span class="chip ${p.mandate.status === "discretionary" ? "gold" : ""}">${esc(p.mandate.status)}${p.mandate.citation ? ` · ${esc(p.mandate.citation)}` : ""}</span>` : ""}${p.strategic_linkage ? `<span class="chip ${none ? "grey" : p.alignment === "direct" ? "gold" : "maroon"}">${p.alignment && !none ? esc(p.alignment) + " · " : ""}${esc(p.strategic_linkage)}</span>` : ""}</div></div>
          <div class="cost"><div class="big">${money(p.cost)}</div><div class="lbl">${Math.round(p.share * 100)}% of this budget · ${Number(p.fte || 0).toFixed(2)} FTE</div>${alloc > 0 && direct > 0 ? `<div class="lbl">${money(direct)} direct + ${money(alloc)} allocated</div>` : ""}${strip}</div></div>
        ${p.beneficiary || p.contributing_divisions ? `<div class="meta">${p.beneficiary ? `<div><span class="g">Beneficiary · </span>${esc(p.beneficiary)}</div>` : ""}${p.contributing_divisions ? `<div><span class="g">Contributing divisions · </span>${esc(p.contributing_divisions)}</div>` : ""}</div>` : ""}
        <div class="measures three">${cell(rk[0], rk[1], m.result)}${cell("Volume", "how much was delivered", m.volume)}${cell("Unit cost", "cost ÷ volume", m.unit_cost)}</div>
        <div class="twenty"><div><span class="lbl" style="color:${C.maroon}">At −20%</span><div>${esc(p.twenty_percent?.minus || "—")}</div></div><div><span class="lbl" style="color:${C.gold}">At +20%</span><div>${esc(p.twenty_percent?.plus || "—")}</div></div></div>${pack}${lines}</section>`;
    }).join("\n");
    const elsewhere = result.accountabilities_elsewhere?.length ? `<section class="card un"><div class="lbl">Accountabilities funded elsewhere</div><p class="note">Real responsibilities of this division whose cost sits outside this budget; noted, not costed here.</p><table>${result.accountabilities_elsewhere.map((a) => `<tr><td style="font-family:Georgia,serif;font-size:14px;white-space:nowrap">${esc(a.name)}</td><td class="gold" style="white-space:nowrap">${esc(a.where)}</td><td class="g">${esc(a.note || "")}</td></tr>`).join("")}</table></section>` : "";
    const checkBlock = `<section class="card plain"><div class="lbl">Budget office check · ${checks.filter((c) => c.ok).length} of ${checks.length} pass</div><div class="checks">${checks.map((c) => `<div><b style="color:${c.ok ? "#3D6B3A" : C.maroon}">${c.ok ? "✓" : "✗"}</b> <span${c.ok ? ' class="g"' : ""}>${esc(c.text)}</span></div>`).join("")}</div></section>`;
    const index = `<table class="glance"><thead><tr><th>Program</th><th>Serves</th><th>Delivers</th><th>So that</th><th class="r">Cost</th><th class="r">FTE</th><th>Mandate</th><th>Alignment</th></tr></thead><tbody>${sorted.map((p) => { const st = parseStatement(p.statement); const none = p.alignment === "none" || /none identified/i.test(p.strategic_linkage || "");
      return `<tr><td><div class="pn">${esc(p.name)}</div><div class="sub" style="color:${p.facing === "internal" ? "#2E7D6E" : C.maroon}">${p.facing === "internal" ? "internal · performance" : "external · outcomes"}</div></td><td><b>${esc(p.beneficiary || st.beneficiary || "—")}</b></td><td>${esc(st.service || "—")}</td><td><b>${esc(st.result || "—")}</b></td><td class="r">${money(p.cost)}<div class="bar" style="width:72px;margin:4px 0 0 auto"><span style="width:${(p.share / (sorted[0].share || 1)) * 100}%;background:${p.is_admin ? C.grey : C.maroon}"></span></div><div class="sub">${Math.round(p.share * 100)}%</div></td><td class="r">${Number(p.fte || 0).toFixed(2)}</td><td><span class="tag" style="border-color:${p.mandate?.status === "discretionary" ? C.gold : C.line};color:${p.mandate?.status === "discretionary" ? C.gold : C.ink}">${esc(p.mandate?.status || "—")}</span>${p.mandate?.citation ? `<div class="sub cite">${esc(p.mandate.citation)}</div>` : ""}</td><td><span class="tag" style="border-color:${none ? C.line : p.alignment === "direct" ? C.gold : C.maroon};color:${none ? C.grey : p.alignment === "direct" ? C.gold : C.maroon}">${none ? "not aligned" : esc(p.alignment || "—")}</span>${!none && p.strategic_linkage ? `<div class="sub cite">${esc(p.strategic_linkage)}</div>` : ""}</td></tr>`; }).join("")}</tbody></table>`;
    const un = unassigned.length ? `<section class="card un"><h2>Unassigned · ${money(unassigned.reduce((s, i) => s + i.amt, 0))}</h2><table>${unassigned.map((l) => `<tr><td class="g">${esc(l.code)}</td><td>${esc(l.desc)} <span class="g">· ${esc(l.cc)}</span></td><td class="r">${money(l.amt)}</td></tr>`).join("")}</table></section>` : "";
    const tail = `<div class="two">${result.cautions?.length ? `<section class="card"><div class="lbl">Where this is a judgment call</div>${result.cautions.map((c) => `<p>${esc(c)}</p>`).join("")}</section>` : ""}
      ${result.data_requests?.length ? `<section class="card"><div class="lbl">What to ask the division for</div><ol>${result.data_requests.map((c) => `<li>${esc(c)}</li>`).join("")}</ol></section>` : ""}</div>`;
    return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(division)} — proposed program layer</title>
<style>
  @page { margin: 16mm 14mm; } body { margin: 0; background: #fff; color: ${C.ink}; font: 15px/1.55 Georgia, "Times New Roman", serif; }
  .wrap { max-width: 860px; margin: 0 auto; padding: 32px 28px 60px; }
  .mast { border-bottom: 3px solid ${C.maroon}; padding-bottom: 12px; margin-bottom: 22px; }
  .kick, .lbl, .g, .src, .comp, .note, .mh, td.g { font-family: Arial, Helvetica, sans-serif; color: ${C.grey}; } .kick { font-size: 13px; } .lbl { font-size: 12px; }
  h1 { font-weight: normal; font-size: 30px; margin: 4px 0 6px; line-height: 1.15; } h2 { font-weight: normal; font-size: 22px; margin: 0 0 6px; line-height: 1.2; }
  .mast p { margin: 0; color: ${C.grey}; font-size: 14px; }
  .cover { display: flex; height: 18px; background: ${C.line}; margin: 10px 0 14px; } .cover span { display: block; border-right: 1px solid #fff; }
  .row { display: grid; grid-template-columns: 1fr 2fr auto; gap: 14px; align-items: center; margin: 5px 0; font-size: 14px; } .bar { height: 9px; background: ${C.paper}; } .bar span { display: block; height: 100%; }
  .r { text-align: right; white-space: nowrap; } .gold { color: ${C.gold}; }
  .legend { font: 12px Arial, Helvetica, sans-serif; color: ${C.grey}; margin: 14px 0 26px; display: flex; gap: 16px; }
  .dot { display: inline-block; width: 9px; height: 9px; border-radius: 9px; margin-right: 6px; vertical-align: middle; } .dot.likely { background: #3D6B3A; } .dot.possible { background: ${C.gold}; } .dot.unlikely { background: ${C.maroon}; }
  .card { border: 1px solid ${C.line}; border-top: 4px solid ${C.maroon}; margin-bottom: 18px; break-inside: avoid; page-break-inside: avoid; } .card.un { border-top-color: ${C.gold}; padding: 16px 20px; }
  .head { display: grid; grid-template-columns: 1fr auto; gap: 20px; padding: 16px 20px 12px; } .head p { margin: 0; max-width: 520px; }
  .cost { text-align: right; min-width: 180px; } .big { font-size: 24px; line-height: 1.1; }
  .strip { display: flex; height: 8px; width: 180px; margin: 10px 0 0 auto; background: ${C.paper}; } .strip span { display: block; }
  .comp { font-size: 11px; margin-top: 4px; display: flex; justify-content: flex-end; gap: 10px; } .comp i { display: inline-block; width: 8px; height: 8px; margin-right: 4px; }
  .chip { display: inline-block; margin-top: 8px; font: 12px Arial, Helvetica, sans-serif; color: ${C.maroonDark}; background: #F6E7B8; padding: 3px 8px; }
  .note { font-size: 13px; margin: 0 0 10px; max-width: 640px; } .note.pad { padding: 0 20px 10px; }
  .measures { display: grid; grid-template-columns: 1fr 1fr; border-top: 1px solid ${C.line}; } .measures > div { padding: 12px 20px 14px; } .measures > div:first-child { border-right: 1px solid ${C.line}; }
  .mh { font-size: 12px; margin-bottom: 8px; } .mh b { font-weight: normal; color: ${C.maroon}; font-size: 13px; margin-right: 6px; }
  .m { display: grid; grid-template-columns: 14px 1fr; gap: 6px; margin-bottom: 9px; font-size: 14px; } .m .dot { margin-top: 6px; } .src { font-size: 12px; }
  .dq { border-top: 1px solid ${C.line}; background: ${C.paper}; padding: 12px 20px; display: grid; grid-template-columns: auto 1fr; gap: 14px; align-items: baseline; } .dq span { font: 12px Arial, Helvetica, sans-serif; color: ${C.gold}; white-space: nowrap; } .dq em { font-size: 16px; }
  .build { border-top: 1px solid ${C.line}; padding: 12px 20px 14px; }
  table { width: 100%; border-collapse: collapse; font: 12.5px Arial, Helvetica, sans-serif; } td { padding: 4px 8px 4px 0; border-top: 1px solid ${C.line}; vertical-align: top; } td.r { padding-right: 0; }
  .two { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; } .two .card { padding: 14px 20px; font-size: 14px; } .two p { margin: 0 0 6px; } ol { margin: 0; padding-left: 20px; }
  .foot { margin-top: 24px; font: 12px Arial, Helvetica, sans-serif; color: ${C.grey}; }
  .stmt { font-size: 15px; } .chips { margin-top: 8px; display: flex; gap: 6px; flex-wrap: wrap; } .chips .chip { margin: 0; background: none; border: 1px solid ${C.grey}; color: ${C.grey}; } .chips .chip.gold { border-color: ${C.gold}; color: ${C.gold}; } .chips .chip.maroon { border-color: ${C.maroon}; color: ${C.maroon}; } .chips .chip.grey { border-color: ${C.line}; }
  .meta { padding: 0 20px 12px; display: flex; gap: 24px; font: 13px Arial, Helvetica, sans-serif; } .card.admin { border-top-color: ${C.grey}; }
  .measures.three { grid-template-columns: 1fr 1fr 1fr; } .measures.three > div { border-right: 1px solid ${C.line}; } .measures.three > div:last-child { border-right: none; }
  .twenty { border-top: 1px solid ${C.line}; background: ${C.paper}; padding: 12px 20px; display: grid; grid-template-columns: 1fr 1fr; gap: 20px; font-size: 14px; } .twenty .lbl { display: block; margin-bottom: 3px; }
  .chips .chip.teal { border-color: #2E7D6E; color: #2E7D6E; } .pack { border-top: 1px solid ${C.line}; padding: 8px 20px; display: flex; gap: 14px; font: 11px Arial, Helvetica, sans-serif; }
  .glance { width: 100%; border-collapse: collapse; font: 11.5px/1.4 Arial, Helvetica, sans-serif; margin: 6px 0 4px; } .glance th { font-weight: normal; color: ${C.grey}; font-size: 10.5px; text-align: left; padding: 4px 6px 6px 0; border-bottom: 1px solid ${C.line}; } .glance td { padding: 7px 6px 7px 0; border-bottom: 1px solid ${C.line}; vertical-align: top; } .glance .pn { font: 13px Georgia, serif; } .glance .sub { font-size: 10px; color: ${C.grey}; margin-top: 2px; } .glance .sub.cite { margin-top: 5px; font-style: italic; } .glance .tag { display: inline-block; font-size: 10.5px; padding: 2px 6px; border: 1px solid; border-radius: 3px; white-space: nowrap; } .glance .r { text-align: right; white-space: nowrap; } .glance .bar { height: 5px; }
  .card.plain { padding: 14px 20px; } .checks { display: grid; grid-template-columns: 1fr 1fr; gap: 5px 20px; font: 13px Arial, Helvetica, sans-serif; }
  @media print { .wrap { padding: 0; max-width: none; } }
</style></head><body><div class="wrap">
  <div class="mast"><div class="kick">Harding · strategic budgeting assistant · Line of Sight</div><h1>${esc(division || "Division")}: proposed program layer</h1>
    <p>Pre-decisional draft to the County Program Standard · ${today} · ${money(total)} across ${programs.length} programs and ${programs.reduce((a, p) => a + (Number(p.fte) || 0), 0).toFixed(1)} FTE · ${Math.round(coverage * 100)}% of budget mapped · ${esc(facingLabel(result.orientation))}</p></div>
  <div class="cover">${sorted.filter((p) => p.cost > 0).map((p, i) => `<span style="width:${Math.max(p.share * 100, 0.4)}%;background:${i % 2 ? C.maroonDark : C.maroon}"></span>`).join("")}${unassigned.length ? `<span style="flex:1;background:${C.gold}"></span>` : ""}</div>
  <p>${esc(result.division_read)}</p>
  ${index}
  <div class="legend"><span>Result measures are outcomes for external programs and service levels for internal ones.</span></div><div class="legend"><span>Data readiness:</span><span>${dot("likely")}exists today</span><span>${dot("possible")}could be assembled</span><span>${dot("unlikely")}would need new collection</span></div>
  ${cards}${elsewhere}${un}${checkBlock}${tail}
  <div class="foot">Dollars are summed from the FY adopted budget export; the grouping, shares, and measures are proposed for discussion with the division. Public Works LLC with Funkhouser &amp; Associates for Frederick County.</div>
</div></body></html>`;
  }
  async function exportHtml() {
    const h = buildHtml(); setHtmlText(h);
    try { await navigator.clipboard.writeText(h); } catch {}
  }

  function buildMarkdown() {
    const today = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
    const L = [];
    L.push(`# ${division || "Division"} — proposed program layer`);
    L.push(`Pre-decisional draft · ${today} · ${money(total)} across ${programs.length} programs · ${Math.round(coverage * 100)}% of budget mapped · ${result.orientation}-facing`);
    L.push("", result.division_read, "", "## Programs at a glance", "", "| Program | Serves | Delivers | So that | Cost | FTE | Mandate | Alignment |", "|---|---|---|---|---:|---:|---|---|");
    sorted.forEach((p) => { const st = parseStatement(p.statement); L.push(`| ${p.name} (${p.facing || ""}) | ${p.beneficiary || st.beneficiary || ""} | ${st.service || ""} | ${st.result || ""} | ${money(p.cost)} (${Math.round(p.share * 100)}%) | ${Number(p.fte || 0).toFixed(2)} | ${p.mandate?.status || ""} | ${p.alignment === "none" ? "not aligned" : p.alignment || ""} |`); });
    sorted.forEach((p, i) => {
      const m = p.measures || {};
      L.push("", `## ${i + 1}. ${p.name}`, "", `_${facingLabel(p.facing || result.orientation)}_`, "", `> ${p.statement}, at ${money(p.cost)} and ${Number(p.fte || 0).toFixed(2)} FTE in FY26.`, "");
      if (p.beneficiary) L.push(`**Beneficiary:** ${p.beneficiary}  `);
      L.push(`**Mandate:** ${p.mandate?.status || "—"}${p.mandate?.citation ? ` — ${p.mandate.citation}` : ""}  `);
      L.push(`**Priority alignment:** ${p.alignment ? p.alignment + " — " : ""}${p.strategic_linkage || "none identified"}  `);
      if (p.contributing_divisions) L.push(`**Contributing divisions:** ${p.contributing_divisions}  `);
      const comp = composition(p.lines).filter((c) => c.amt > 0);
      if (comp.length) L.push(`**Composition:** ${comp.map((c) => `${c.label} ${Math.round((c.amt / p.cost) * 100)}%`).join(" · ")}`);
      L.push("", "| Measure | | Source | Data |", "|---|---|---|---|");
      [[resultKind(p.facing || result.orientation)[0], m.result], ["Volume", m.volume], ["Unit cost", m.unit_cost]].forEach(([k, x]) => L.push(`| ${k} | ${x?.metric || "—"} | ${x?.source || ""} | ${x?.exists_today || ""} |`));
      L.push("", `**Evidence pack:** ${packItems(p).map(([k, ok]) => `${ok ? "●" : "○"} ${k}`).join("  ")}`);
      L.push("", `**At −20%:** ${p.twenty_percent?.minus || "—"}  `, `**At +20%:** ${p.twenty_percent?.plus || "—"}`);
      if (p.lines.length) {
        L.push("", "**How the cost is built**", "");
        if (p.allocation_note) L.push(p.allocation_note, "");
        L.push("| Costing center | Account | Line | Share | Allocated |", "|---|---|---|---:|---:|");
        p.lines.forEach((l) => L.push(`| ${l.cc} | ${l.code} | ${l.desc} | ${Math.round(l.share * 100)}% | ${money(l.alloc)} |`));
      }
    });
    if (result.accountabilities_elsewhere?.length) {
      L.push("", "## Accountabilities funded elsewhere", "", "| Accountability | Where the money sits | Note |", "|---|---|---|");
      result.accountabilities_elsewhere.forEach((a) => L.push(`| ${a.name} | ${a.where} | ${a.note || ""} |`));
    }
    if (unassigned.length) {
      L.push("", "## Unassigned", "", "| Costing center | Account | Line | Amount |", "|---|---|---|---:|");
      unassigned.forEach((l) => L.push(`| ${l.cc} | ${l.code} | ${l.desc} | ${money(l.amt)} |`));
    }
    L.push("", "## Budget office check", ""); checks.forEach((c) => L.push(`- ${c.ok ? "✓" : "✗"} ${c.text}`));
    if (result.cautions?.length) { L.push("", "## Where this is a judgment call", ""); result.cautions.forEach((c) => L.push(`- ${c}`)); }
    if (result.data_requests?.length) { L.push("", "## What to ask the division for", ""); result.data_requests.forEach((c, i) => L.push(`${i + 1}. ${c}`)); }
    return L.join("\n");
  }
  async function exportMarkdown() {
    const md = buildMarkdown(); setMdText(md);
    try { await navigator.clipboard.writeText(md); } catch {}
  }

  const ready = items.length > 0 && status !== "thinking";

  // ---- atoms
  const Label = ({ children }) => <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginBottom: 4 }}>{children}</div>;
  const Select = ({ value, onChange, allowNone }) => (
    <select value={value} onChange={(e) => onChange(e.target.value)}
      style={{ width: "100%", padding: "8px 10px", border: `1px solid ${C.line}`, background: C.white, fontFamily: sans, fontSize: 13, color: C.ink }}>
      {allowNone && <option value="">(none)</option>}
      {headers.map((h) => <option key={h} value={h}>{h}</option>)}
    </select>
  );
  const Drop = ({ onChange, accept, text, name }) => (
    <label style={{ display: "block", padding: 14, border: `1px dashed ${name ? C.line : C.maroon}`, borderRadius: 4, background: C.paper, cursor: "pointer", fontFamily: sans, fontSize: 13 }}>
      <input type="file" accept={accept} onChange={onChange} style={{ display: "none" }} />
      {name ? <span>{name}</span> : <span style={{ color: C.grey }}>{text}</span>}
    </label>
  );
  const readiness = { likely: "#3D6B3A", possible: C.gold, unlikely: C.maroon };
  const Dot = ({ v }) => <span style={{ display: "inline-block", width: 9, height: 9, borderRadius: 9, background: readiness[v] || C.grey, marginRight: 6, verticalAlign: "middle" }} />;
  const MeasureCell = ({ kind, sub, m, last }) => (
    <div style={{ padding: `12px ${pad}px 14px`, borderRight: last ? "none" : `1px solid ${C.line}`, borderBottom: narrow ? `1px solid ${C.line}` : "none" }}>
      <div style={{ fontFamily: sans, fontSize: 12, marginBottom: 6 }}><span style={{ color: C.maroon }}>{kind}</span> <span style={{ color: C.grey }}>{sub}</span></div>
      {!m?.metric ? <div style={{ fontFamily: sans, fontSize: 13, color: C.maroon }}>Not proposed.</div> : (
        <div style={{ display: "grid", gridTemplateColumns: "14px 1fr", gap: 6, alignItems: "start" }}>
          <div style={{ paddingTop: 5 }}><Dot v={m.exists_today} /></div>
          <div><div style={{ fontSize: 14, lineHeight: 1.45 }}>{m.metric}</div><div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginTop: 2 }}>{m.source}</div></div>
        </div>
      )}
    </div>
  );
  const parseStatement = (st) => {
    const t = String(st || "").replace(/\s+/g, " ").trim();
    const m = t.match(/^(.*?)\s+(?:delivers|provides|offers)\s+(.*?)\s+to\s+(.*?)\s+so that\s+(.*?)(?:,\s*at\s+\$.*)?$/i);
    if (m) return { service: m[2], beneficiary: m[3], result: m[4].replace(/[.]+$/, "") };
    const st2 = t.match(/so that\s+(.*?)(?:,\s*at\s+\$.*)?$/i);
    return { service: "", beneficiary: "", result: st2 ? st2[1].replace(/[.]+$/, "") : "" };
  };
  const facingLabel = (f) => (f === "internal" ? "Internal · performance-based" : f === "external" ? "External · outcomes-based" : "Mixed · both");
  const resultKind = (f) => (f === "internal" ? ["Service level", "timeliness, reliability, quality — what the division controls"] : ["Outcome", "the community result this program contributes to"]);
  const packItems = (p) => [
    ["Cost", p.cost > 0], ["Reach", !!p.beneficiary], ["Necessity", !!(p.twenty_percent?.minus)], ["Performance", !!(p.measures?.volume?.metric && p.measures?.unit_cost?.metric)],
    ["Impact", p.measures?.result?.metric && p.measures?.result?.exists_today !== "unlikely"], ["Statutory", !!p.mandate?.status],
  ];
  const composition = (lines) => {
    const cats = [
      { label: "Staff", color: C.maroon, test: (c) => c.startsWith("50") },
      { label: "Contracts & services", color: C.gold, test: (c) => c.startsWith("52") || c.startsWith("54") },
      { label: "Software, supplies & equipment", color: "#9A7B2E", test: (c) => c.startsWith("51") || c.startsWith("53") },
      { label: "Other", color: C.grey, test: () => true },
    ].map((c) => ({ ...c, amt: 0 }));
    lines.forEach((l) => { const c = cats.find((k) => k.test(l.code || "")); c.amt += l.alloc ?? l.amt; });
    return cats;
  };
  const LineTable = ({ lines, showShare }) => (
    <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 13 }}>
      <tbody>
        {lines.map((l) => (
          <tr key={l.id} style={{ borderTop: `1px solid ${C.line}` }}>
            {!narrow && <td style={{ padding: "5px 8px 5px 0", color: C.grey, width: 70 }}>{l.code}</td>}
            <td style={{ padding: "5px 8px 5px 0" }}>{l.desc}{l.cc && l.cc !== l.desc && <span style={{ color: C.grey }}> · {l.cc}</span>}{narrow && showShare && l.share < 1 && <div style={{ color: C.grey, fontSize: 12 }}>{Math.round(l.share * 100)}% of {money(l.amt)}</div>}</td>
            {showShare && !narrow && <td style={{ padding: "5px 8px", color: C.grey, textAlign: "right", whiteSpace: "nowrap" }}>{l.share < 1 ? `${Math.round(l.share * 100)}% of ${money(l.amt)}` : ""}</td>}
            <td style={{ padding: "5px 0", textAlign: "right", whiteSpace: "nowrap" }}>{money(l.alloc ?? l.amt)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const Btn = ({ primary, children, className, ...rest }) => (
    <button {...rest} className={`btn ${primary ? "btn-primary" : "btn-secondary"} ${className || ""}`}>{children}</button>
  );
  const STAGES = [["reading", "Reading the budget"], ["thinking", "Thinking it through"], ["drafting", "Drafting programs"], ["finishing", "Checking the draft"]];
  const stageIdx = STAGES.findIndex(([k]) => k === stage);
  const renderRunStatus = () => (
    <div className="card" style={{ padding: `${narrow ? 14 : 16}px ${pad}px`, marginBottom: 20 }}>
      <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
        <div style={{ fontSize: narrow ? 16 : 18 }}>{progress || "Reading the budget…"}</div>
        <div style={{ fontFamily: sans, fontSize: 12, color: C.grey }}>{items.length} lines · {new Set(items.map((i) => i.cc)).size} costing centers · sent as {sendCount}</div>
      </div>
      <div className="indet" style={{ marginTop: 12 }}><span /></div>
      <div style={{ display: "grid", gridTemplateColumns: `repeat(${STAGES.length}, 1fr)`, gap: 6, marginTop: 10 }}>
        {STAGES.map(([k, label], i) => {
          const state = i < stageIdx ? "done" : i === stageIdx ? "active" : "todo";
          return (
            <div key={k} style={{ fontFamily: sans, fontSize: narrow ? 10.5 : 11.5, color: state === "todo" ? C.line : state === "active" ? C.maroon : C.grey, display: "flex", alignItems: "center", gap: 6, minWidth: 0 }}>
              <span style={{ width: 8, height: 8, borderRadius: 8, background: state === "todo" ? C.line : state === "active" ? C.maroon : C.grey, flexShrink: 0 }} />
              <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{label}{k === "drafting" && draft?.programs?.length ? ` · ${draft.programs.length}` : ""}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
  const renderSkeleton = () => (
    <section className="card program-card skel" style={{ marginBottom: 20 }}>
      <div style={{ padding: `18px ${pad}px 14px`, display: "grid", gridTemplateColumns: narrow ? "1fr" : "1fr auto", gap: 20 }}>
        <div><div className="sk" style={{ width: 90, height: 10, marginBottom: 10 }} /><div className="sk" style={{ width: "55%", height: 22, marginBottom: 10 }} /><div className="sk" style={{ width: "92%", height: 12, marginBottom: 6 }} /><div className="sk" style={{ width: "70%", height: 12 }} /></div>
        {!narrow && <div style={{ textAlign: "right" }}><div className="sk" style={{ width: 120, height: 26, marginLeft: "auto", marginBottom: 8 }} /><div className="sk" style={{ width: 150, height: 10, marginLeft: "auto" }} /></div>}
      </div>
      <div style={{ borderTop: `1px solid ${C.line}`, display: "grid", gridTemplateColumns: narrow ? "1fr" : "1fr 1fr 1fr" }}>
        {[0, 1, 2].map((i) => <div key={i} style={{ padding: `12px ${pad}px` }}><div className="sk" style={{ width: 70, height: 9, marginBottom: 8 }} /><div className="sk" style={{ width: "85%", height: 12 }} /></div>)}
      </div>
    </section>
  );

  return (
    <div style={{ minHeight: "100vh", background: C.sand, color: C.ink, fontFamily: serif }}>
      <style>{`
        * { -webkit-tap-highlight-color: transparent; }
        .card { background: ${C.white}; border: 1px solid ${C.line}; border-radius: 6px; box-shadow: 0 1px 2px rgba(44,44,44,0.05); }
        .program-card { border-top: 4px solid ${C.maroon}; }
        .btn { padding: 10px 16px; border-radius: 4px; font: 13.5px Arial, Helvetica, sans-serif; cursor: pointer; transition: background .15s, color .15s, border-color .15s, opacity .15s; }
        .btn-primary { background: ${C.maroon}; color: ${C.sand}; border: 1px solid ${C.maroon}; } .btn-primary:hover:not(:disabled) { background: ${C.maroonDark}; }
        .btn-primary:disabled { background: ${C.line}; border-color: ${C.line}; color: ${C.grey}; cursor: default; }
        .btn-secondary { background: transparent; color: ${C.maroon}; border: 1px solid ${C.maroon}; } .btn-secondary:hover { background: rgba(107,26,42,0.06); }
        .btn-ghost { background: transparent; color: ${C.sand}; border: 1px solid rgba(244,238,226,0.45); border-radius: 4px; font: 12px Arial, Helvetica, sans-serif; cursor: pointer; } .btn-ghost:hover { border-color: ${C.sand}; }
        .fade { animation: fadeUp .35s ease-out both; } @keyframes fadeUp { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: none; } }
        .indet { height: 4px; background: ${C.paper}; border-radius: 4px; overflow: hidden; position: relative; }
        .indet span { position: absolute; left: -40%; top: 0; bottom: 0; width: 40%; background: linear-gradient(90deg, transparent, ${C.gold}, transparent); animation: slide 1.4s ease-in-out infinite; }
        @keyframes slide { to { left: 100%; } }
        .sk { background: ${C.paper}; border-radius: 3px; animation: pulse 1.3s ease-in-out infinite; } @keyframes pulse { 50% { opacity: .45; } }
        .skel { border-top-color: ${C.line}; }
        .pill { display: inline-flex; align-items: center; gap: 7px; font: 12px Arial, Helvetica, sans-serif; padding: 4px 10px; border-radius: 999px; white-space: nowrap; }
        .spin { width: 10px; height: 10px; border: 2px solid rgba(244,238,226,0.35); border-top-color: ${C.sand}; border-radius: 50%; animation: rot .8s linear infinite; } @keyframes rot { to { transform: rotate(360deg); } }
        @media print {
          @page { margin: 16mm 14mm; }
          body { background: #fff !important; }
          .no-print { display: none !important; }
          .print-layout { display: block !important; padding: 0 !important; max-width: none !important; }
          .appbar { position: static !important; background: #fff !important; color: #2C2C2C !important; border-bottom: 3px solid #6B1A2A; box-shadow: none !important; }
          .program-card { break-inside: avoid; page-break-inside: avoid; box-shadow: none; }
          .cost-build-body { display: block !important; }
          .fade { animation: none; }
          a { color: inherit; }
        }
      `}</style>

      {/* app bar */}
      <div className="appbar" style={{ position: "sticky", top: 0, zIndex: 20, background: C.maroon, color: C.sand, boxShadow: "0 1px 0 rgba(0,0,0,0.15)" }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", padding: narrow ? "10px 16px" : "12px 32px", display: "flex", alignItems: "center", gap: 14, minHeight: 48 }}>
          <div title="Harding — after the Budget and Accounting Act of 1921" style={{ width: 28, height: 28, borderRadius: 6, background: C.gold, display: "grid", placeItems: "center", flexShrink: 0 }}>
            <svg width="20" height="20" viewBox="0 0 20 20" aria-label="abacus">
              <rect x="1.5" y="1.5" width="17" height="17" rx="1.5" fill="none" stroke={C.maroon} strokeWidth="1.6" />
              <line x1="1.5" y1="6" x2="18.5" y2="6" stroke={C.maroon} strokeWidth="1.1" />
              <line x1="1.5" y1="10" x2="18.5" y2="10" stroke={C.maroon} strokeWidth="1.1" />
              <line x1="1.5" y1="14" x2="18.5" y2="14" stroke={C.maroon} strokeWidth="1.1" />
              <circle cx="5" cy="6" r="1.6" fill={C.maroon} /><circle cx="8.4" cy="6" r="1.6" fill={C.maroon} /><circle cx="15" cy="6" r="1.6" fill={C.maroon} />
              <circle cx="5" cy="10" r="1.6" fill={C.maroon} /><circle cx="12" cy="10" r="1.6" fill={C.maroon} /><circle cx="15.4" cy="10" r="1.6" fill={C.maroon} />
              <circle cx="5" cy="14" r="1.6" fill={C.maroon} /><circle cx="8.4" cy="14" r="1.6" fill={C.maroon} /><circle cx="11.8" cy="14" r="1.6" fill={C.maroon} />
            </svg>
          </div>
          <div style={{ minWidth: 0, flex: 1 }}>
            <div style={{ display: "flex", alignItems: "baseline", gap: 10, whiteSpace: "nowrap", overflow: "hidden" }}>
              <span style={{ fontFamily: serif, fontSize: narrow ? 17 : 19, letterSpacing: 0.3 }}>Harding</span>
              {!narrow && <span style={{ fontFamily: sans, fontSize: 11.5, opacity: 0.75, letterSpacing: 0.4, textTransform: "uppercase" }}>Strategic budgeting assistant</span>}
            </div>
            {(division || fileName) && <div style={{ fontFamily: sans, fontSize: 11.5, opacity: 0.8, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{division || fileName}</div>}
            {narrow && !(division || fileName) && <div style={{ fontFamily: sans, fontSize: 11, opacity: 0.75, letterSpacing: 0.4, textTransform: "uppercase" }}>Strategic budgeting assistant</div>}
          </div>
          {status === "thinking" && <span className="pill" style={{ background: "rgba(0,0,0,0.18)" }}><span className="spin" />{narrow ? "Working" : STAGES[Math.max(stageIdx, 0)][1]}</span>}
          {status === "done" && !narrow && <span className="pill" style={{ background: "rgba(201,156,28,0.22)", color: "#F6E7B8" }}>Draft ready</span>}
          {narrow && result && <button className="btn-ghost" style={{ padding: "6px 10px" }} onClick={() => setSetupOpen((v) => !v)}>{setupOpen ? "Hide setup" : "Setup"}</button>}
          {!narrow && <Btn primary onClick={analyze} disabled={!ready} style={{ padding: "8px 14px" }}>{status === "thinking" ? "Working…" : result ? "Run again" : "Propose programs"}</Btn>}
        </div>
      </div>

      <div className="print-layout" style={{ maxWidth: 1100, margin: "0 auto", padding: narrow ? "16px 16px 48px" : "24px 32px 64px", display: "grid", gridTemplateColumns: narrow ? "1fr" : "300px 1fr", gap: narrow ? 20 : 32, alignItems: "start" }}>
        {(setupOpen || !narrow) && <aside className="no-print card" style={{ padding: `${narrow ? 14 : 18}px ${pad}px`, display: "flex", flexDirection: "column", gap: 16, position: narrow ? "static" : "sticky", top: 72 }}>
          <div style={{ fontFamily: sans, fontSize: 11, letterSpacing: 0.8, textTransform: "uppercase", color: C.grey }}>Setup</div>
          <div><Label>Budget export (.csv or .xlsx)</Label>
            <Drop onChange={onBudget} accept=".csv,.xlsx,.xls" name={fileName} text="Choose the division's line-item budget" /></div>

          {headers.length > 0 && (
            <details style={{ fontFamily: sans, fontSize: 13 }}>
              <summary style={{ cursor: "pointer", color: C.grey }}>Column mapping</summary>
              <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 12 }}>
                <div><Label>Costing center name</Label><Select value={cols.cc} onChange={(v) => setCols({ ...cols, cc: v })} allowNone /></div>
                <div><Label>GL account name</Label><Select value={cols.gl} onChange={(v) => setCols({ ...cols, gl: v })} allowNone /></div>
                <div><Label>GL account number</Label><Select value={cols.code} onChange={(v) => setCols({ ...cols, code: v })} allowNone /></div>
                <div><Label>Adopted budget</Label><Select value={cols.amt} onChange={(v) => setCols({ ...cols, amt: v })} /></div>
              </div>
            </details>
          )}

          <div><Label>Division</Label>
            <input value={division} onChange={(e) => setDivision(e.target.value)} placeholder="Filled from the file when present"
              style={{ width: "100%", padding: "8px 10px", border: `1px solid ${C.line}`, background: C.white, fontFamily: sans, fontSize: 13, color: C.ink, boxSizing: "border-box" }} /></div>

          {items.length > 0 && (
            <div style={{ fontFamily: sans, fontSize: 13, color: C.grey, lineHeight: 1.5 }}>
              {items.length} funded lines · {money(total)}<br />
              Personnel is {Math.round(personnelShare * 100)}% of the total, roughly {headcount} staff{personnelShare > 0.8 ? "; programs will be costed by staff time" : ""}.
              {zeroCount > 0 && <><br />{zeroCount} zero, revenue, or recovery lines set aside{recoveryTotal ? `, including ${money(recoveryTotal)} of cost recoveries billed to other units` : ""}.</>}
            </div>
          )}

          {narrow && <Btn primary onClick={analyze} disabled={!ready} style={{ padding: "12px 16px", fontSize: 14 }}>{status === "thinking" ? "Working…" : result ? "Run again" : "Propose programs"}</Btn>}
          {error && <div style={{ fontFamily: sans, fontSize: 13, color: C.maroon, lineHeight: 1.5, borderLeft: `3px solid ${C.maroon}`, paddingLeft: 10 }}>{error}</div>}
        </aside>}

        <main style={{ minWidth: 0 }}>
          {!live && status !== "thinking" && (
            <div className="card fade" style={{ padding: `${narrow ? 18 : 24}px ${pad}px` }}>
              <div style={{ fontFamily: sans, fontSize: 11, letterSpacing: 0.8, textTransform: "uppercase", color: C.grey, marginBottom: 12 }}>How it works</div>
              <div style={{ display: "grid", gridTemplateColumns: narrow ? "1fr" : "1fr 1fr 1fr", gap: narrow ? 14 : 22 }}>
                {[["1", "Upload a budget export", "The county's standard line-item file. Revenue, zero, and recovery lines are set aside; personnel and small lines are pooled per costing center."],
                  ["2", "Programs are drafted to the Standard", "One statement each, full cost and FTE, three measures of three kinds, mandate, priority alignment, and the twenty-percent question — cards appear as each one lands."],
                  ["3", "The budget office's checks run", "Sums to appropriation, span, one administration program, no effort verbs, measures from real systems. Then export to PDF, HTML, Markdown, or CSV."]].map(([n, t, d]) => (
                  <div key={n} style={{ display: "grid", gridTemplateColumns: "28px 1fr", gap: 10 }}>
                    <div style={{ width: 26, height: 26, borderRadius: 13, background: C.maroon, color: C.sand, fontFamily: sans, fontSize: 12, display: "grid", placeItems: "center" }}>{n}</div>
                    <div><div style={{ fontSize: 15.5, marginBottom: 4 }}>{t}</div><div style={{ fontFamily: sans, fontSize: 12.5, color: C.grey, lineHeight: 1.5 }}>{d}</div></div>
                  </div>
                ))}
              </div>
              <p style={{ margin: "18px 0 0", fontSize: 14, color: C.grey, lineHeight: 1.55, maxWidth: 640 }}>
                What comes out is a first draft meant to be argued with in a room with the division director. Dollars are summed from your file; only the grouping, shares, and language are proposed.
              </p>
            </div>
          )}

          {status === "thinking" && renderRunStatus()}

          {live && (
            <div>
              {/* overview: stats, coverage bar, one-line read, program index */}
              {live.division_read && !programs.length && <p style={{ fontSize: 15.5, lineHeight: 1.6, maxWidth: 640, margin: "0 0 20px" }}>{live.division_read}</p>}
              {programs.length > 0 && <div className="card" style={{ padding: `${narrow ? 14 : 18}px ${pad}px`, marginBottom: 20 }}>
                <div style={{ display: "grid", gridTemplateColumns: narrow ? "1fr 1fr" : "repeat(4, 1fr)", gap: narrow ? 10 : 16, marginBottom: 14 }}>
                  {[[`${programs.length}`, "programs"], [money(total), "in this budget"], [programs.reduce((a, p) => a + (Number(p.fte) || 0), 0).toFixed(1), "FTE"], [`${Math.round(coverage * 100)}%`, streaming ? "mapped so far" : "mapped"]].map(([v, l]) => (
                    <div key={l}><div style={{ fontSize: narrow ? 20 : 24, lineHeight: 1.1 }}>{v}</div><div style={{ fontFamily: sans, fontSize: 11.5, color: C.grey, marginTop: 3 }}>{l}</div></div>
                  ))}
                </div>
                <div style={{ display: "flex", height: 14, width: "100%", background: C.paper, borderRadius: 3, overflow: "hidden" }}>
                  {sorted.filter((p) => p.cost > 0).map((p, i) => (
                    <div key={p.name} title={`${p.name}: ${money(p.cost)}`}
                      style={{ width: `${Math.max(p.share * 100, 0.4)}%`, background: p.is_admin ? C.grey : i % 2 === 0 ? C.maroon : C.maroonDark, borderRight: `1px solid ${C.white}`, transition: "width .4s" }} />
                  ))}
                  {!streaming && unassigned.length > 0 && <div style={{ flex: 1, background: C.gold }} title="Unassigned" />}
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", fontFamily: sans, fontSize: 11.5, color: C.grey, marginTop: 6, gap: 10 }}>
                  <span>{facingLabel(live.orientation)}</span><span style={{ textAlign: "right" }}>{streaming ? "shares settle when the last program lands" : coverage > 0.995 ? "every dollar sits in a program" : `${money(unassigned.reduce((s, i) => s + i.amt, 0))} unassigned`}</span>
                </div>
                <p style={{ fontSize: 15, lineHeight: 1.6, maxWidth: 640, margin: "16px 0 16px" }}>{live.division_read}</p>
                {!narrow ? (
                  <div style={{ overflowX: "auto", margin: "0 -4px" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 12.5, lineHeight: 1.45 }}>
                      <thead>
                        <tr style={{ color: C.grey, fontSize: 11, textAlign: "left" }}>
                          {["Program", "Serves", "Delivers", "So that", "Cost", "FTE", "Mandate", "Alignment"].map((h, i) => (
                            <th key={h} style={{ fontWeight: "normal", padding: "6px 8px 8px 4px", borderBottom: `1px solid ${C.line}`, textAlign: i === 4 || i === 5 ? "right" : "left", whiteSpace: "nowrap" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sorted.map((p) => { const st = parseStatement(p.statement); const none = p.alignment === "none" || /none identified/i.test(p.strategic_linkage || "");
                          return (
                          <tr key={p.name} style={{ borderBottom: `1px solid ${C.line}`, verticalAlign: "top" }}>
                            <td style={{ padding: "8px 8px 8px 4px", minWidth: 150 }}>
                              <a href={`#prog-${encodeURIComponent(p.name)}`} style={{ color: C.ink, textDecoration: "none", fontFamily: serif, fontSize: 14 }}>{p.name}</a>
                              <div style={{ fontSize: 10.5, color: p.facing === "internal" ? "#2E7D6E" : C.maroon, marginTop: 2 }}>{p.facing === "internal" ? "internal · performance" : "external · outcomes"}</div>
                            </td>
                            <td style={{ padding: "8px 8px 8px 4px", minWidth: 140, color: C.ink, fontWeight: "bold" }}>{p.beneficiary || st.beneficiary || "—"}</td>
                            <td style={{ padding: "8px 8px 8px 4px", minWidth: 150 }}>{st.service || "—"}</td>
                            <td style={{ padding: "8px 8px 8px 4px", minWidth: 170, fontWeight: "bold" }}>{st.result || "—"}</td>
                            <td style={{ padding: "8px 8px 8px 4px", textAlign: "right", whiteSpace: "nowrap" }}>
                              <div>{money(p.cost)}</div>
                              <div style={{ height: 5, width: 72, background: C.paper, marginLeft: "auto", marginTop: 4, position: "relative" }}><div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: `${(p.share / (sorted[0].share || 1)) * 100}%`, background: p.is_admin ? C.grey : C.maroon }} /></div>
                              <div style={{ fontSize: 10.5, color: C.grey, marginTop: 2 }}>{Math.round(p.share * 100)}%</div>
                            </td>
                            <td style={{ padding: "8px 8px 8px 4px", textAlign: "right", whiteSpace: "nowrap" }}>{Number(p.fte || 0).toFixed(2)}</td>
                            <td style={{ padding: "8px 8px 8px 4px", minWidth: 120 }}>
                              <span style={{ display: "inline-block", fontSize: 11, padding: "2px 7px", border: `1px solid ${p.mandate?.status === "discretionary" ? C.gold : C.line}`, borderRadius: 3, color: p.mandate?.status === "discretionary" ? C.gold : C.ink, whiteSpace: "nowrap" }}>{p.mandate?.status || "—"}</span>
                              {p.mandate?.citation && <div style={{ fontSize: 10.5, color: C.grey, marginTop: 5, fontStyle: "italic", lineHeight: 1.35 }}>{p.mandate.citation}</div>}
                            </td>
                            <td style={{ padding: "8px 4px 8px 4px", minWidth: 130 }}>
                              <span style={{ display: "inline-block", fontSize: 11, padding: "2px 7px", border: `1px solid ${none ? C.line : p.alignment === "direct" ? C.gold : C.maroon}`, borderRadius: 3, color: none ? C.grey : p.alignment === "direct" ? C.gold : C.maroon, whiteSpace: "nowrap" }}>{none ? "not aligned" : p.alignment || "—"}</span>
                              {!none && p.strategic_linkage && <div style={{ fontSize: 10.5, color: C.grey, marginTop: 5, fontStyle: "italic", lineHeight: 1.35 }}>{p.strategic_linkage}</div>}
                            </td>
                          </tr>);
                        })}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {sorted.map((p) => { const st = parseStatement(p.statement); const none = p.alignment === "none" || /none identified/i.test(p.strategic_linkage || "");
                      return (
                      <a key={p.name} href={`#prog-${encodeURIComponent(p.name)}`} style={{ color: C.ink, textDecoration: "none", borderTop: `1px solid ${C.line}`, paddingTop: 8 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "baseline" }}>
                          <span style={{ fontFamily: serif, fontSize: 15 }}>{p.name}</span>
                          <span style={{ fontFamily: sans, fontSize: 12, color: C.grey, whiteSpace: "nowrap" }}>{money(p.cost)} · {Math.round(p.share * 100)}% · {Number(p.fte || 0).toFixed(1)} FTE</span>
                        </div>
                        <div style={{ height: 5, background: C.paper, position: "relative", margin: "5px 0 6px" }}><div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: `${(p.share / (sorted[0].share || 1)) * 100}%`, background: p.is_admin ? C.grey : C.maroon }} /></div>
                        <div style={{ fontFamily: sans, fontSize: 12, lineHeight: 1.5, color: C.grey }}>
                          <div><span style={{ color: C.grey }}>Serves </span><span style={{ color: C.ink, fontWeight: "bold" }}>{p.beneficiary || st.beneficiary || "—"}</span></div>
                          {st.result && <div><span style={{ color: C.grey }}>So that </span><span style={{ color: C.ink, fontWeight: "bold" }}>{st.result}</span></div>}
                          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 6 }}>
                            <span style={{ fontSize: 10.5, padding: "1px 6px", border: `1px solid ${p.mandate?.status === "discretionary" ? C.gold : C.line}`, borderRadius: 3, color: p.mandate?.status === "discretionary" ? C.gold : C.ink }}>{p.mandate?.status || "—"}</span>
                            <span style={{ fontSize: 10.5, padding: "1px 6px", border: `1px solid ${none ? C.line : p.alignment === "direct" ? C.gold : C.maroon}`, borderRadius: 3, color: none ? C.grey : p.alignment === "direct" ? C.gold : C.maroon }}>{none ? "not aligned" : p.alignment || "—"}</span>
                            <span style={{ fontSize: 10.5, padding: "1px 6px", border: `1px solid ${C.line}`, borderRadius: 3, color: p.facing === "internal" ? "#2E7D6E" : C.maroon }}>{p.facing === "internal" ? "internal" : "external"}</span>
                          </div>
                        </div>
                      </a>);
                    })}
                  </div>
                )}
                <div style={{ marginTop: 14, fontFamily: sans, fontSize: 12, color: C.grey, display: "flex", gap: 18, flexWrap: "wrap" }}>
                  <span>Data readiness:</span>
                  <span><Dot v="likely" /> exists today</span>
                  <span><Dot v="possible" /> could be assembled</span>
                  <span><Dot v="unlikely" /> would need new collection</span>
                  <span style={{ marginLeft: "auto" }}>Result measures: outcome for external programs, service level for internal.</span>
                </div>
              </div>}

              {/* program cards */}
              {sorted.map((p, idx) => {
                const key = p.name; const isOpen = open[key];
                const comp = composition(p.lines);
                const m = p.measures || {};
                const mandateColor = p.mandate?.status === "discretionary" ? C.gold : C.grey;
                return (
                  <section key={key} id={`prog-${encodeURIComponent(p.name)}`} className="card program-card fade"
                    style={{ borderTopColor: p.is_admin ? C.grey : C.maroon, marginBottom: 20, scrollMarginTop: 72, overflow: "hidden" }}>
                    {/* header */}
                    <div style={{ padding: `18px ${pad}px 14px`, display: "grid", gridTemplateColumns: narrow ? "1fr" : "1fr auto", gap: narrow ? 12 : 20, alignItems: "start" }}>
                      <div>
                        <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginBottom: 4 }}>Program {idx + 1}{streaming ? "" : ` of ${sorted.length}`}{p.is_admin ? " · administration" : ""}{streaming ? " · draft" : ""}</div>
                        <h2 style={{ fontWeight: "normal", fontSize: narrow ? 20 : 23, margin: "0 0 8px", lineHeight: 1.2 }}>{p.name}</h2>
                        <p style={{ margin: 0, fontSize: 15.5, lineHeight: 1.55, maxWidth: 580 }}>
                          {p.statement}{p.statement && !/[.!?]$/.test(p.statement) ? "" : ""}, at {money(p.cost)} and {Number(p.fte || 0).toFixed(2)} FTE in FY26.
                        </p>
                        <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                          {p.facing && <span style={{ fontFamily: sans, fontSize: 11, color: p.facing === "internal" ? "#2E7D6E" : C.maroon, border: `1px solid ${p.facing === "internal" ? "#2E7D6E" : C.maroon}`, padding: "2px 8px" }}>{facingLabel(p.facing)}</span>}
                          {p.mandate?.status && <span title={p.mandate.citation || ""} style={{ fontFamily: sans, fontSize: 11, color: mandateColor, border: `1px solid ${mandateColor}`, padding: "2px 8px" }}>{p.mandate.status}{p.mandate.citation ? ` · ${p.mandate.citation}` : ""}</span>}
                          {p.strategic_linkage && (() => { const none = p.alignment === "none" || /none identified/i.test(p.strategic_linkage); const col = none ? C.grey : p.alignment === "direct" ? C.gold : C.maroon;
                            return <span style={{ fontFamily: sans, fontSize: 11, color: col, border: `1px solid ${none ? C.line : col}`, padding: "2px 8px" }}>{p.alignment && !none ? `${p.alignment} · ` : ""}{none && p.alignment === "none" && !/none identified/i.test(p.strategic_linkage) ? "not aligned · " : ""}{p.strategic_linkage}</span>; })()}
                        </div>
                      </div>
                      <div style={{ textAlign: narrow ? "left" : "right", minWidth: narrow ? 0 : 180, borderTop: narrow ? `1px solid ${C.line}` : "none", paddingTop: narrow ? 10 : 0 }}>
                        <div style={{ fontSize: narrow ? 22 : 26, lineHeight: 1.1 }}>{money(p.cost)}</div>
                        <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginTop: 4 }}>{Math.round(p.share * 100)}% of this budget · {Number(p.fte || 0).toFixed(2)} FTE</div>
                        {(() => { const direct = p.lines.filter((l) => l.share >= 1).reduce((a, l) => a + l.alloc, 0); const alloc = p.cost - direct;
                          return alloc > 0 && direct > 0 ? <div style={{ fontFamily: sans, fontSize: 11, color: C.grey, marginTop: 2 }}>{money(direct)} direct + {money(alloc)} allocated</div> : alloc > 0 ? <div style={{ fontFamily: sans, fontSize: 11, color: C.grey, marginTop: 2 }}>fully allocated from shared lines</div> : null; })()}
                        <div style={{ marginTop: 10 }}>
                          <div style={{ display: "flex", height: 8, width: narrow ? "100%" : 180, marginLeft: narrow ? 0 : "auto", background: C.paper }}>
                            {comp.map((c) => c.amt > 0 && <div key={c.label} title={`${c.label}: ${money(c.amt)}`} style={{ width: `${(c.amt / p.cost) * 100}%`, background: c.color }} />)}
                          </div>
                          <div style={{ fontFamily: sans, fontSize: 11, color: C.grey, marginTop: 4, display: "flex", justifyContent: narrow ? "flex-start" : "flex-end", gap: 10, flexWrap: "wrap" }}>
                            {comp.filter((c) => c.amt > 0).map((c) => <span key={c.label}><span style={{ display: "inline-block", width: 8, height: 8, background: c.color, marginRight: 4, verticalAlign: "middle" }} />{c.label} {Math.round((c.amt / p.cost) * 100)}%</span>)}
                          </div>
                        </div>
                      </div>
                    </div>
                    {(p.beneficiary || p.contributing_divisions) && (
                      <div style={{ padding: `0 ${pad}px 14px`, display: "grid", gridTemplateColumns: p.contributing_divisions && !narrow ? "1fr 1fr" : "1fr", gap: narrow ? 8 : 20, fontFamily: sans, fontSize: 13, lineHeight: 1.5 }}>
                        {p.beneficiary && <div><span style={{ color: C.grey }}>Beneficiary · </span>{p.beneficiary}</div>}
                        {p.contributing_divisions && <div><span style={{ color: C.grey }}>Contributing divisions · </span>{p.contributing_divisions}</div>}
                      </div>
                    )}

                    {/* three measures */}
                    <div style={{ borderTop: `1px solid ${C.line}`, display: "grid", gridTemplateColumns: narrow ? "1fr" : "1fr 1fr 1fr" }}>
                      <MeasureCell kind={resultKind(p.facing || live.orientation)[0]} sub={resultKind(p.facing || live.orientation)[1]} m={m.result} last={narrow} />
                      <MeasureCell kind="Volume" sub="how much was delivered" m={m.volume} last={narrow} />
                      <MeasureCell kind="Unit cost" sub="cost ÷ volume" m={m.unit_cost} last />
                    </div>

                    {/* twenty percent question */}
                    <div style={{ borderTop: narrow ? "none" : `1px solid ${C.line}`, background: C.paper, padding: `14px ${pad}px`, display: "grid", gridTemplateColumns: narrow ? "1fr" : "1fr 1fr", gap: narrow ? 12 : 22 }}>
                      <div><div style={{ fontFamily: sans, fontSize: 12, color: C.maroon, marginBottom: 4 }}>At −20%</div><div style={{ fontSize: 14.5, lineHeight: 1.5 }}>{p.twenty_percent?.minus || "—"}</div></div>
                      <div><div style={{ fontFamily: sans, fontSize: 12, color: C.gold, marginBottom: 4 }}>At +20%</div><div style={{ fontSize: 14.5, lineHeight: 1.5 }}>{p.twenty_percent?.plus || "—"}</div></div>
                    </div>

                    {/* evidence pack readiness */}
                    <div style={{ borderTop: `1px solid ${C.line}`, padding: `10px ${pad}px`, fontFamily: sans, fontSize: 11, color: C.grey, display: "flex", gap: narrow ? 10 : 16, flexWrap: "wrap", alignItems: "center" }}>
                      <span>Evidence pack</span>
                      {packItems(p).map(([k, ok]) => <span key={k} style={{ color: ok ? C.ink : C.maroon }}><span style={{ color: ok ? "#3D6B3A" : C.maroon, marginRight: 4 }}>{ok ? "●" : "○"}</span>{k}</span>)}
                    </div>

                    {/* cost build */}
                    {p.lines.length > 0 && (
                      <div style={{ borderTop: `1px solid ${C.line}`, padding: `12px ${pad}px 16px` }}>
                        <button className="no-print" onClick={() => setOpen({ ...open, [key]: !open[key] })}
                          style={{ background: "none", border: "none", padding: 0, fontFamily: sans, fontSize: 13, color: C.maroon, cursor: "pointer" }}>
                          {isOpen ? "▾" : "▸"} How {money(p.cost)} is built · {p.lines.length} lines{p.lines.some((l) => l.share < 1) ? ", some shared" : ""}
                        </button>
                        {isOpen && (
                          <div className="cost-build-body" style={{ marginTop: 12 }}>
                            {p.allocation_note && <p style={{ fontFamily: sans, fontSize: 13, color: C.grey, margin: "0 0 10px", maxWidth: 640, lineHeight: 1.5 }}>{p.allocation_note}</p>}
                            <LineTable lines={p.lines} showShare />
                          </div>
                        )}
                      </div>
                    )}
                  </section>
                );
              })}

              {streaming && renderSkeleton()}

              {result?.accountabilities_elsewhere?.length > 0 && (
                <section className="card fade" style={{ borderTop: `4px solid ${C.gold}`, marginBottom: 20, padding: `16px ${pad}px` }}>
                  <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginBottom: 4 }}>Accountabilities funded elsewhere</div>
                  <p style={{ margin: "0 0 10px", fontSize: 14, color: C.grey, maxWidth: 640, lineHeight: 1.5 }}>Real responsibilities of this division whose cost sits outside this budget. Under the standard they are noted, not costed here.</p>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: sans, fontSize: 13 }}><tbody>
                    {result.accountabilities_elsewhere.map((a, i) => (
                      <tr key={i} style={{ borderTop: `1px solid ${C.line}` }}>
                        <td style={{ padding: "6px 10px 6px 0", fontFamily: serif, fontSize: 14.5 }}>{a.name}{narrow && <div style={{ fontFamily: sans, fontSize: 12, color: C.gold }}>{a.where}</div>}</td>
                        {!narrow && <td style={{ padding: "6px 10px 6px 0", color: C.gold, whiteSpace: "nowrap" }}>{a.where}</td>}
                        <td style={{ padding: "6px 0", color: C.grey }}>{a.note}</td>
                      </tr>
                    ))}
                  </tbody></table>
                </section>
              )}

              {result && <section className="card fade" style={{ marginBottom: 20, padding: `16px ${pad}px` }}>
                <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginBottom: 8 }}>Budget office check · {checks.filter((c) => c.ok).length} of {checks.length} pass</div>
                <div style={{ display: "grid", gridTemplateColumns: narrow ? "1fr" : "1fr 1fr", columnGap: 24, rowGap: 6, fontFamily: sans, fontSize: 13, lineHeight: 1.45 }}>
                  {checks.map((c, i) => (
                    <div key={i} style={{ display: "grid", gridTemplateColumns: "14px 1fr", gap: 8 }}>
                      <span style={{ color: c.ok ? "#3D6B3A" : C.maroon, fontWeight: "bold" }}>{c.ok ? "✓" : "✗"}</span><span style={{ color: c.ok ? C.grey : C.ink }}>{c.text}</span>
                    </div>
                  ))}
                </div>
              </section>}

              {result && unassigned.length > 0 && (
                <section className="card fade" style={{ borderTop: `4px solid ${C.gold}`, marginBottom: 20, padding: `18px ${pad}px` }}>
                  <h2 style={{ fontWeight: "normal", fontSize: 21, margin: "0 0 6px" }}>Unassigned · {money(unassigned.reduce((s, i) => s + i.amt, 0))}</h2>
                  <p style={{ fontFamily: sans, fontSize: 13, color: C.grey, margin: "0 0 10px" }}>These need a home before the map is usable.</p>
                  <LineTable lines={unassigned} />
                </section>
              )}

              {result && (result.cautions?.length > 0 || result.data_requests?.length > 0) && (
                <div style={{ display: "grid", gridTemplateColumns: result.cautions?.length && result.data_requests?.length && !narrow ? "1fr 1fr" : "1fr", gap: 20, marginBottom: 20 }}>
                  {result.cautions?.length > 0 && (
                    <section className="card fade" style={{ padding: `16px ${pad}px`, fontSize: 14, lineHeight: 1.6 }}>
                      <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginBottom: 8 }}>Where this is a judgment call</div>
                      {result.cautions.map((c, i) => <p key={i} style={{ margin: "0 0 8px" }}>{c}</p>)}
                    </section>
                  )}
                  {result.data_requests?.length > 0 && (
                    <section className="card fade" style={{ padding: `16px ${pad}px`, fontSize: 14, lineHeight: 1.6 }}>
                      <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginBottom: 8 }}>What to ask the division for</div>
                      <ol style={{ margin: 0, paddingLeft: 20 }}>{result.data_requests.map((c, i) => <li key={i} style={{ marginBottom: 6 }}>{c}</li>)}</ol>
                    </section>
                  )}
                </div>
              )}

              {result && <div className="no-print" style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
                <Btn primary onClick={exportHtml}>{htmlText ? "Report copied as HTML" : "Copy report as HTML (for PDF)"}</Btn>
                <Btn onClick={exportMarkdown}>{mdText ? "Report copied as Markdown" : "Copy report as Markdown"}</Btn>
                <Btn onClick={exportCsv}>{copied ? "Copied to clipboard" : "Copy allocation (CSV)"}</Btn>
              </div>}
              {result && <div style={{ marginTop: 28, paddingTop: 14, borderTop: `1px solid ${C.line}`, fontFamily: sans, fontSize: 12, color: C.grey, lineHeight: 1.6, maxWidth: 720 }}>
                Pre-decisional draft to the County Program Standard · {new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" })}. Dollars are summed from the FY adopted budget export; the grouping, shares, FTE, and measures are proposed for discussion with the division. Public Works LLC with Funkhouser &amp; Associates for Frederick County.
              </div>}
              {htmlText && (
                <div className="no-print" style={{ marginTop: 16, borderLeft: `3px solid ${C.gold}`, paddingLeft: 14 }}>
                  <div style={{ fontFamily: sans, fontSize: 13, color: C.ink, lineHeight: 1.6 }}>
                    The report is on your clipboard as a self-contained web page. Paste it into any text editor, save it as <b>report.html</b>, open that file in your browser, and choose Print → Save as PDF. It's already laid out for paper.
                  </div>
                  <textarea readOnly value={htmlText} onFocus={(e) => e.target.select()} rows={4}
                    style={{ width: "100%", boxSizing: "border-box", marginTop: 8, padding: 10, border: `1px solid ${C.line}`, background: C.white, fontFamily: "monospace", fontSize: 11, color: C.grey }} />
                </div>
              )}
              {mdText && (
                <div className="no-print" style={{ marginTop: 16 }}>
                  <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginBottom: 6 }}>Full report as Markdown — paste into a doc, or hand it to Claude to produce a styled PDF.</div>
                  <textarea readOnly value={mdText} onFocus={(e) => e.target.select()} rows={8}
                    style={{ width: "100%", boxSizing: "border-box", padding: 10, border: `1px solid ${C.line}`, background: C.white, fontFamily: "monospace", fontSize: 12, color: C.ink }} />
                </div>
              )}
              {csvText && (
                <div className="no-print" style={{ marginTop: 16 }}>
                  <div style={{ fontFamily: sans, fontSize: 12, color: C.grey, marginBottom: 6 }}>One row per line item per program, with the share and allocated amount. Paste into a blank sheet.</div>
                  <textarea readOnly value={csvText} onFocus={(e) => e.target.select()} rows={10}
                    style={{ width: "100%", boxSizing: "border-box", padding: 10, border: `1px solid ${C.line}`, background: C.white, fontFamily: "monospace", fontSize: 12, color: C.ink }} />
                </div>
              )}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
