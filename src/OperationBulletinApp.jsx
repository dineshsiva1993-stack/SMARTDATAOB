import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Plus, X, Users, GitBranch, Gauge, ClipboardList, Settings, TrendingUp, Boxes,
  Search, Download, Trash2, ChevronDown, ChevronRight, Layers, FolderOpen, Eye, EyeOff, Home, LogOut,
  Save, Upload, FileSpreadsheet, Check, Menu as MenuIcon, Pencil, Printer, Share2,
} from "lucide-react";
import { BarChart, Bar, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from "recharts";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

const GSD_LOGO_DATA_URI = "data:image/jpeg;base64,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";


// Login gate for Operation Bulletin. This seeds the very first admin user the first time the
// app runs on a device — after that, all users (including this one) live in localStorage
// under OB_USERS_KEY and are managed from the in-app Admin screen.
const OB_CREDENTIALS = [{ user: "admin", pass: "smartdataob@2026" }];
const OB_SESSION_KEY = "obAuthed";
const OB_SESSION_USER_KEY = "obAuthedUser";
const OB_USERS_KEY = "ob-users";

// Menus that can be individually granted to a non-admin user from the Admin screen.
const OB_MENU_OPTIONS = [
  { key: "dashboard", label: "Dashboard" },
  { key: "newstyle", label: "+ New Style" },
  { key: "target", label: "Target & Manpower" },
  { key: "ops", label: "Operations" },
  { key: "subentries", label: "Sub Entries" },
  { key: "balance", label: "SMV & Balance" },
  { key: "rampup", label: "Ramp-up" },
  { key: "templateupload", label: "OB Template Upload" },
  { key: "savedob", label: "Saved OB (Buyer-wise)" },
];

// Reads the user list from localStorage; on first-ever run (no users saved yet) it seeds a
// single full-access Admin user from OB_CREDENTIALS above so there's always a way in.
function loadObUsersRaw() {
  try {
    const raw = localStorage.getItem(OB_USERS_KEY);
    const parsed = raw ? JSON.parse(raw) : null;
    if (Array.isArray(parsed) && parsed.length) return parsed;
  } catch (e) {}
  return OB_CREDENTIALS.map((c) => ({
    id: uid(),
    username: c.user,
    password: c.pass,
    isAdmin: true,
    menus: OB_MENU_OPTIONS.map((m) => m.key),
    factoryIds: [],
  }));
}

// Simple persisted-users store, same pattern as useGsdLibrary below.
function useObUsers() {
  const [users, setUsersState] = useState(loadObUsersRaw);
  const save = (next) => {
    setUsersState(next);
    try { localStorage.setItem(OB_USERS_KEY, JSON.stringify(next)); } catch (e) {}
  };
  return { users, save };
}

// =====================================================================================
// Standalone Operation Bulletin app — split out of SGPS so it can be used on its own.
// Reads/writes the SAME localStorage keys as the main SGPS app ("sgps-styles",
// "sgps-factories", "sgps-current-factory"), so any style created in SGPS shows up here
// automatically, and any Operation Bulletin you build here is visible back in SGPS.
// =====================================================================================

const uid = () => Math.random().toString(36).slice(2, 10);
const money = (n) => (isFinite(n) ? n.toLocaleString("en-IN", { maximumFractionDigits: 2 }) : "0");
const num = (v) => parseFloat(v) || 0;
// Single source of truth for "how many people are on this operation" — always derived fresh from
// Op + Hp (Sewing rows use these two split boxes) rather than trusting the op.manpower string,
// which can fall out of sync with Op/Hp if a row was inserted/imported through a path that didn't
// go through updateOpHpQc. Falls back to the manpower field (Cutting/QC/Packing rows, which only
// have a single Manpower box) and finally to 1, same as the old `num(op.manpower || 1)` default.
const allocatedMpOf = (op) => {
  const opHp = num(op?.opCount) + num(op?.hpCount);
  if (opHp > 0) return opHp;
  return num(op?.manpower) || 1;
};
const SHIFT_MINUTES = 480; // standard 8-hour shift, used for efficiency & capacity calc

const DEPT_KEYS = [
  { key: "cutting", label: "Cutting", entryType: "cutting" },
  { key: "sewing", label: "Sewing", entryType: "sewing" },
  { key: "packing", label: "Packing", entryType: "packing" },
];
const blankDepartments = () => DEPT_KEYS.map((d) => ({ key: d.key, smv: "", manpower: "", machines: "" }));

// ---------------- Operation Bulletin (OB) ----------------
// One row per operation (loading, shoulder attach, neck binding...). Each operation belongs to a
// section (Cutting / Sewing / QC / Packing) mirroring a real factory Operation Breakdown sheet —
// SL NO, Operation, Machine, SAM, Manpower — used to build section-wise SMV and Line Balancing.
const OB_SECTIONS = [
  { key: "cutting", label: "Cutting" },
  { key: "sewing", label: "Sewing" },
  { key: "qc", label: "QC" },
  { key: "packing", label: "Packing" },
];
const blankOBOperation = () => ({ id: uid(), slNo: "", name: "", section: "sewing", machine: "", sam: "", manpower: "1", opCount: "", hpCount: "", qcCount: "", isHeading: false });

// Section-wise SMV (sum of SAM per section) + combined totals matching a real OB sheet:
// SMV (Sew to QC) = Sewing + QC, SMV (Sew to Pack) = Sewing + QC + Packing, Total SMV = all sections incl Cutting.
function computeOBSummary(operations) {
  const ops = (operations || []).filter((op) => !op.isHeading);
  const byKey = Object.fromEntries(OB_SECTIONS.map((s) => [s.key, { smv: 0, count: 0, manpower: 0 }]));
  ops.forEach((op) => {
    const s = byKey[op.section] || byKey.sewing;
    s.smv += num(op.sam);
    s.count += 1;
    s.manpower += allocatedMpOf(op);
  });
  const sewToQcSmv = byKey.sewing.smv + byKey.qc.smv;
  const sewToPackSmv = sewToQcSmv + byKey.packing.smv;
  const totalSmv = sewToPackSmv + byKey.cutting.smv;
  return { sections: byKey, sewToQcSmv, sewToPackSmv, totalSmv };
}

// Line balancing — per-operation hourly capacity (manpower × 60 / SAM) for the Sewing section only,
// since that's the line whose stations must be balanced against each other. Bottleneck = the
// operation with the lowest capacity (it caps what the whole line can produce per hour).
function computeLineBalancing(operations) {
  const sewOps = (operations || []).filter((op) => op.section === "sewing" && num(op.sam) > 0);
  const rows = sewOps.map((op) => ({
    ...op,
    capacityPerHr: allocatedMpOf(op) > 0 ? (allocatedMpOf(op) * 60) / num(op.sam) : 0,
  })).filter((r) => r.capacityPerHr > 0);
  if (rows.length === 0) return { rows: [], bottleneck: null, avgCapacity: 0, ucl: 0, lcl: 0 };
  const sorted = [...rows].sort((a, b) => a.capacityPerHr - b.capacityPerHr);
  const bottleneck = sorted[0];
  const avgCapacity = rows.reduce((a, r) => a + r.capacityPerHr, 0) / rows.length;
  const ucl = Math.max(...rows.map((r) => r.capacityPerHr));
  const lcl = bottleneck.capacityPerHr;
  return { rows, bottleneck, avgCapacity, ucl, lcl };
}

// Suggest a standard garment-line machine from an operation's name (e.g. "Shoulder overlock" → Overlock).
// Purely a suggestion — the Machine field stays a free-text manual entry so this can always be overridden.
// Standard Tirupur garment-line machine codes (matches factory Operation Breakdown sheets) —
// used both as quick-pick suggestions (datalist) and for keyword-based auto-suggest below.
const OB_MACHINE_LIST = [
  "SNLS", "ZIG ZAG", "PICOTING", "2T O/L", "3T O/L", "4T O/L", "5T O/L",
  "2T F/L", "3T F/L", "5T F/L", "DNLS", "M/N", "B/A", "B/H", "B/T", "HT", "SNAP", "A/P", "SMOKER",
];
// Standard Packing-section process stations — Packing doesn't run on sewing machines, so it gets
// its own quick-pick list (shown in the Machine field's datalist when Section = Packing).
const OB_PACKING_LIST = [
  "ALL",
  "IRONING", "GARMENT SIZE SEPARER", "BARCODE STICKER ON HANG TAG", "HANG TAG USE GUN",
  "ASSORT MAKING", "GARMENT FOLDING & GARMENT INSERT TO POLY BAG", "POLYBAG STICKER",
  "FINAL CHECKING ON POLYBAG & TAG", "METAL DETECTION", "COTTON BOX MAKE AND CLOSE 25 PCS & LOADING",
  "BOX STITCHER 10 NOS",
];
// Standard Cutting-section process stations — Cutting doesn't run on sewing machines either, so it
// gets its own quick-pick list (shown in the Machine field's datalist when Section = Cutting).
const OB_CUTTING_LIST = [
  "ALL", "FABRIC INSPECTION", "LAYERING", "CUTTER", "SORTING", "STICKRING", "BUNDLING",
];
// Manpower role name for each Packing process, so the Manpower box placeholder reflects who's
// doing that job (e.g. Ironing -> "Ironer") instead of the generic word "Manpower".
const OB_PACKING_ROLES = {
  "IRONING": "Ironer",
  "GARMENT SIZE SEPARER": "Size Separator",
  "BARCODE STICKER ON HANG TAG": "Sticker",
  "HANG TAG USE GUN": "Tag Gun Operator",
  "ASSORT MAKING": "Assort Maker",
  "GARMENT FOLDING & GARMENT INSERT TO POLY BAG": "Folder & Packer",
  "POLYBAG STICKER": "Sticker",
  "FINAL CHECKING ON POLYBAG & TAG": "Checker",
  "METAL DETECTION": "Metal Detector Operator",
  "COTTON BOX MAKE AND CLOSE 25 PCS & LOADING": "Box Maker & Loader",
  "BOX STITCHER 10 NOS": "Box Stitcher",
};
function packingRoleLabel(machine) {
  return OB_PACKING_ROLES[(machine || "").trim().toUpperCase()] || "";
}
const MACHINE_SUGGEST_RULES = [
  { keywords: ["picot", "picoting"], machine: "PICOTING" },
  { keywords: ["zigzag", "zig zag"], machine: "ZIG ZAG" },
  { keywords: ["overlock 2", "2 thread overlock"], machine: "2T O/L" },
  { keywords: ["overlock 3", "3 thread overlock"], machine: "3T O/L" },
  { keywords: ["overlock 5", "5 thread overlock", "safety stitch"], machine: "5T O/L" },
  { keywords: ["overlock", "over lock", "serger"], machine: "4T O/L" },
  { keywords: ["flatlock 2", "2 thread flatlock"], machine: "2T F/L" },
  { keywords: ["flatlock 5", "5 thread flatlock"], machine: "5T F/L" },
  { keywords: ["flatlock", "flat lock"], machine: "3T F/L" },
  { keywords: ["button hole", "buttonhole"], machine: "B/H" },
  { keywords: ["button attach", "button"], machine: "B/A" },
  { keywords: ["bartack", "bar tack"], machine: "B/T" },
  { keywords: ["snap", "rivet"], machine: "SNAP" },
  { keywords: ["merrow", "multi needle", "kansai", "chain stitch", "coverstitch", "cover stitch"], machine: "M/N" },
  { keywords: ["double needle", "twin needle"], machine: "DNLS" },
  { keywords: ["single needle", "top stitch", "topstitch", "lockstitch", "feed of the arm", "hemming", "hem"], machine: "SNLS" },
  { keywords: ["heat transfer", "hot fix", "heat seal"], machine: "HT" },
  { keywords: ["auto puller", "attach press", "pressing attach"], machine: "A/P" },
  { keywords: ["smoke", "smoking", "napping"], machine: "SMOKER" },
];
function suggestMachine(opName) {
  const n = (opName || "").toLowerCase().trim();
  if (!n) return "";
  const hit = MACHINE_SUGGEST_RULES.find((r) => r.keywords.some((k) => n.includes(k)));
  return hit ? hit.machine : "";
}

// Manpower balance per operation against a customer/head target efficiency % — mirrors a factory
// Operation Breakdown sheet's "Plan Cap → Man Power → OP/HP → EFF%" columns: how many operators an
// operation NEEDS to hit the target qty/day at the given head efficiency %, versus how many are
// actually allocated (manual entry), so the shortfall/surplus is visible at a glance. Covers every
// section (Cutting, Sewing, QC, Packing) that has a SAM entered — not Sewing only — so the same
// Required MP / Eff% columns show up for Cutting, QC and Packing rows in the OB table too.
// sections: optional array to restrict which OB_SECTIONS keys are included (defaults to all four).
function computeManpowerBalance(operations, targetQtyPerDay, headEffPct, shiftMinutes = SHIFT_MINUTES, sections = null) {
  const target = num(targetQtyPerDay);
  const headEff = num(headEffPct) || 100;
  const hourlyTarget = target > 0 ? target / (shiftMinutes / 60) : 0;
  const rows = (operations || [])
    .filter((op) => !op.isHeading && num(op.sam) > 0 && (!sections || sections.includes(op.section)))
    .map((op) => {
      const sam = num(op.sam);
      const cap100PerHr = 60 / sam; // pieces/hr per operator at 100% efficiency
      const planCapPerHr = cap100PerHr * (headEff / 100);
      const requiredMp = hourlyTarget > 0 && planCapPerHr > 0 ? hourlyTarget / planCapPerHr : 0;
      const allocatedMp = allocatedMpOf(op);
      const balanceMp = allocatedMp - requiredMp;
      const effPct = allocatedMp > 0 ? (requiredMp / allocatedMp) * 100 : 0;
      return { id: op.id, name: op.name, slNo: op.slNo, section: op.section, sam, cap100PerHr, planCapPerHr, requiredMp, allocatedMp, balanceMp, effPct };
    });
  return { hourlyTarget, headEff, rows };
}

// Currencies available for the CM Calculation section — factory CPM/CM figures can be quoted in
// any of these depending on buyer/costing sheet; pick one, every CM stat re-labels with its symbol.
const CM_CURRENCIES = [
  { code: "USD", symbol: "$" },
  { code: "INR", symbol: "₹" },
  { code: "EUR", symbol: "€" },
  { code: "GBP", symbol: "£" },
  { code: "BDT", symbol: "৳" },
  { code: "LKR", symbol: "Rs" },
];
function cmCurrencySymbol(code) {
  return (CM_CURRENCIES.find((c) => c.code === code) || CM_CURRENCIES[0]).symbol;
}

// Day-wise production ramp-up plan. Day 1 onward (up to peak) is manually entered by the planner —
// every line's ramp-up curve is different. The planner enters each day's qty PER LINE (the same
// figure they'd read off one line's ramp-up curve); we multiply by No. of Lines to get that day's
// true total across the whole order. Once the manual entries stop, the rest of the order is
// auto-filled at the Line Target (Customer Target Qty/Day): remaining qty = Order Qty − manually
// entered total so far, and remaining days = that ÷ Line Target (rounded up), so total days to
// complete the order comes out automatically — no manual lead-time entry needed.
function computeRampUpPlan(orderQty, lineTarget, manualDays, noOfLines) {
  const order = num(orderQty);
  const target = num(lineTarget);
  const lines = num(noOfLines) > 0 ? num(noOfLines) : 1;
  const manual = (manualDays || []).map((r, i) => ({ day: i + 1, qty: num(r.qty) * lines, perLineQty: num(r.qty), manual: true }));
  const manualSum = manual.reduce((a, r) => a + r.qty, 0);
  const remaining = Math.max(0, order - manualSum);
  // Exact (fractional) days needed at the Line Target — e.g. 44400 remaining ÷ 2400/day = 18.5 days.
  // Kept alongside the rounded-up day count: the row list needs whole days to enumerate, but the
  // lead-time figure should show the precise fraction (18.5 days), not the rounded-up 19.
  const autoDaysExact = target > 0 && remaining > 0 ? remaining / target : 0;
  const autoDaysCount = Math.ceil(autoDaysExact);
  const autoRows = [];
  let left = remaining;
  for (let i = 0; i < autoDaysCount; i++) {
    const qty = i === autoDaysCount - 1 ? left : target;
    autoRows.push({ day: manual.length + i + 1, qty: Math.round(qty), perLineQty: Math.round(qty) / lines, manual: false });
    left -= target;
  }
  return {
    rows: [...manual, ...autoRows],
    manualSum,
    remaining,
    autoDaysCount,
    autoDaysExact,
    totalDays: manual.length + autoDaysCount,
    totalDaysExact: manual.length + autoDaysExact,
  };
}

// Machine-wise usage summary — how many operations (and how much manpower) sit on each machine type,
// so the planner can see machine loading at a glance.
function computeMachineUsage(operations) {
  const map = {};
  (operations || []).forEach((op) => {
    const m = (op.machine || "").trim();
    if (!m) return;
    if (!map[m]) map[m] = { machine: m, opCount: 0, manpower: 0 };
    map[m].opCount += 1;
    map[m].manpower += allocatedMpOf(op);
  });
  return Object.values(map).sort((a, b) => b.manpower - a.manpower);
}

// Helper summary — total Operators (OP) vs Helpers (HP) tallied from the Manpower Split boxes
// across every operation, so the planner can see the overall operator:helper headcount at a glance.
function computeHelperSummary(operations) {
  const rows = (operations || []).filter((op) => !op.isHeading && (num(op.opCount) > 0 || num(op.hpCount) > 0));
  const totalOp = rows.reduce((a, op) => a + num(op.opCount), 0);
  const totalHp = rows.reduce((a, op) => a + num(op.hpCount), 0);
  return { totalOp, totalHp, total: totalOp + totalHp };
}

// Suggested sewing-line manpower from the customer's target qty/day: (target × sewing SAM) /
// (available minutes/day × target efficiency%) — standard line-planning formula.
function suggestedManpower(targetQtyPerDay, sewingSam, targetEfficiencyPct, shiftMinutes = SHIFT_MINUTES) {
  const target = num(targetQtyPerDay);
  const sam = num(sewingSam);
  const eff = num(targetEfficiencyPct) || 100;
  if (target <= 0 || sam <= 0) return 0;
  return (target * sam) / (shiftMinutes * (eff / 100));
}

// =====================================================================================
// GSD (General Sewing Data) engine — ported from GSD Studio reference.
// A predetermined motion-time system: every operation is built from small timed "elements"
// (a GSD library motion code, a machine-sew stitch, or a manual TMU entry). 1 TMU = 0.0006 min,
// so minutes = TMU / (1000 * 0.0006) = TMU / 600... GSD Studio's reference uses TMU/2000 for its
// minute conversion (its own house standard), which this keeps for parity with that file.
// =====================================================================================
const GSD_CODE_LIBRARY = {
  MG2T: { desc: "MATCH & GET 2 PARTS TOGETHER", tmu: { 0: 76, 5: 53.4, 15: 62.9, 30: 71.3, 45: 81.9, 80: 92.6 } },
  MG2S: { desc: "MATCH & GET 2 PARTS SEPARATELY", tmu: { 0: 107, 5: 63.4, 15: 81.4, 30: 98.2, 45: 117.6, 80: 137.7 } },
  FOOT: { desc: "(MATCH) PARTS TO FOOT", tmu: { 0: 38, 5: 32.5, 15: 37.1, 30: 41.7, 45: 47.6, 80: 53.6 } },
  MAPE: { desc: "MATCH & ADD PART EASY", tmu: { 0: 50, 5: 37.1, 15: 45.2, 30: 53.6, 45: 63.1, 80: 72.8 } },
  MAP1: { desc: "MATCH & ADD PART WITH 1 HAND", tmu: { 0: 56, 5: 44.4, 15: 54.3, 30: 62.2, 45: 72.2, 80: 82.6 } },
  MAP2: { desc: "MATCH & ADD PART WITH 2 HAND", tmu: { 0: 69, 5: 58.7, 15: 68.8, 30: 74.8, 45: 85.0, 80: 96.7 } },
  AM2P: { desc: "ALIGN OR ADJUST (MATCH) 2 PARTS", tmu: { 0: 61, 5: 56.1, 15: 59.9, 30: 63.0, 45: 66.5, 80: 73.0 } },
  AJPT: { desc: "ALIGN OR ADJUST (MATCH) ONE PART", tmu: { 0: 43, 5: 36.5, 15: 41.6, 30: 45.1, 45: 49.8, 80: 54.8 } },
  ARPN: { desc: "ALIGN AND REPOSITION PARTS UNDER FOOT", tmu: { 0: 75, 5: 70.2, 15: 74.0, 30: 77.1, 45: 80.6, 80: 87.1 } },
  APSH: { desc: "ALIGN/ASIDE — PUSH AWAY (SLIDING)", tmu: { 0: 24, 5: 17.0, 15: 20.6, 30: 25.2, 45: 29.3, 80: 33.8 } },
  FFLD: { desc: "FROM FOLD", tmu: { 0: 43, 5: 37.9, 15: 41.7, 30: 44.8, 45: 48.3, 80: 54.8 } },
  FCRS: { desc: "FROM CREASE", tmu: { 0: 28, 5: 21.0, 15: 24.6, 30: 29.2, 45: 33.3, 80: 37.8 } },
  FUNF: { desc: "FROM UNFOLD OR LAY OUT", tmu: { 0: 23, 5: 18.0, 15: 21.8, 30: 24.9, 45: 28.4, 80: 34.9 } },
  TCUT: { desc: "TRIM CUT WITH SCISSORS", tmu: { 0: 50, 5: 37.7, 15: 47.4, 30: 55.6, 45: 65.6, 80: 75.9 } },
  TCAT: { desc: "TRIM CUT WITH SCISSORS — ADDITIONAL", tmu: { 0: 25, 5: 19.1, 15: 23.9, 30: 28.3, 45: 33.6, 80: 39.0 } },
  TBLD: { desc: "TRIM CUT THREAD WITH FIXED BLADE", tmu: { 0: 33, 5: 27.9, 15: 31.7, 30: 34.8, 45: 38.3, 80: 44.8 } },
  TDCH: { desc: "TRIM / DE-CHAIN PARTS WITH SCISSORS", tmu: { 0: 49 } },
  AS1H: { desc: "ASIDE PART WITH 1 HAND", tmu: { 0: 23, 5: 17.2, 15: 21.2, 30: 26.2, 45: 32.0, 80: 38.0 } },
  AS2H: { desc: "ASIDE PART WITH 2 HAND", tmu: { 0: 42, 5: 31.2, 15: 35.6, 30: 40.2, 45: 46.0, 80: 52.0 } },
  MS1A: { desc: "MACHINE SEW 1 CM — APPROXIMATELY", tmu: { 0: 17 } },
  MS1B: { desc: "MACHINE SEW 1 CM — ACCURATE (WITHIN 1 CM)", tmu: { 0: 26 } },
  MS1C: { desc: "MACHINE SEW 1 CM — PRECISE (WITHIN 1/2 CM)", tmu: { 0: 37 } },
  MHDW: { desc: "MACHINE HAND WHEEL TO RAISE/LOWER NEEDLE", tmu: { 0: 46 } },
  MBTB: { desc: "MACHINE BACK TACK AT BEGINNING (LEVER)", tmu: { 0: 34 } },
  MBTE: { desc: "MACHINE BACK TACK AT END (LEVER)", tmu: { 0: 37 } },
  MABT: { desc: "MACHINE BACK TACK (AUTOMATIC)", tmu: { 0: 10 } },
  MBBT: { desc: "MACHINE BACK TACK (BUTTON)", tmu: { 0: 24 } },
  MFBT: { desc: "MACHINE BACK TACK (FOOT)", tmu: { 0: 18 } },
  GP1E: { desc: "GET PART WITH 1 HAND (EASY)", tmu: { 0: 14, 5: 6.8, 15: 10.1, 30: 14.1, 45: 18.3, 80: 22.6 } },
  GP1H: { desc: "GET PART WITH 1 HAND", tmu: { 0: 20, 5: 14.1, 15: 19.2, 30: 22.7, 45: 27.4, 80: 32.4 } },
  GP2H: { desc: "GET PART WITH 2 HAND", tmu: { 0: 33, 5: 27.1, 15: 32.2, 30: 35.7, 45: 40.4, 80: 45.4 } },
  GPCO: { desc: "GET PART CONTACT ONLY", tmu: { 0: 9, 5: 2.8, 15: 6.1, 30: 9.5, 45: 13.5, 80: 17.3 } },
  GPOH: { desc: "GET PART OTHER HAND ONLY", tmu: { 0: 6 } },
  GPAG: { desc: "GET PART BY ADJUSTING GRASP", tmu: { 0: 10, 5: 8.0, 15: 11.0, 30: 14.2, 45: 18.0, 80: 22.1 } },
  PPAL: { desc: "PUT PART TO APPROXIMATE LOCATION", tmu: { 0: 10, 5: 3.0, 15: 6.6, 30: 11.2, 45: 15.3, 80: 19.8 } },
  PPOH: { desc: "PUT PART TO OTHER HAND", tmu: { 0: 6, 5: 2.5, 15: 5.5, 30: 10.0, 45: 15.5, 80: 22.2 } },
  PPST: { desc: "PUT PARTS TO STACK", tmu: { 0: 14, 5: 3.2, 15: 7.6, 30: 12.2, 45: 18.0, 80: 24.0 } },
  PPL1: { desc: "PUT PART, LOCATE ONCE (1)", tmu: { 0: 27, 5: 21.1, 15: 25.7, 30: 30.3, 45: 36.3, 80: 41.4 } },
  PPL2: { desc: "PUT PART, LOCATE TWICE (2)", tmu: { 0: 47, 5: 40.2, 15: 45.8, 30: 50.4, 45: 55.4, 80: 61.5 } },
  F: { desc: "FOOT OR SHORT LEG MOTION", tmu: { 0: 9 } },
  P: { desc: "PACE OR STEP TO MOVE BODY", tmu: { 0: 18 } },
  B: { desc: "BEND AND ARISE", tmu: { 0: 61 } },
  BD: { desc: "BEND DOWN", tmu: { 0: 29 } },
  AB: { desc: "ARISE FROM BEND", tmu: { 0: 32 } },
  SIT: { desc: "SIT DOWN IN CHAIR", tmu: { 0: 35 } },
  STD: { desc: "STAND FROM CHAIR", tmu: { 0: 44 } },
  E: { desc: "EYE ACTION (SIMPLE BINARY CHECK)", tmu: { 0: 7 } },
  C: { desc: "CRANK", tmu: { 0: 15 } },
  R: { desc: "RE-GRASP", tmu: { 0: 6 } },
  A: { desc: "APPLY PRESSURE", tmu: { 0: 14 } },
};
const GSD_TENSION = { N: { label: "Nil", factor: 1 }, L: { label: "Low", factor: 1.1 }, M: { label: "Medium", factor: 1.2 }, H: { label: "High", factor: 1.4 } };
const GSD_STOP = { A: { label: "> 1 cm", add: 0 }, B: { label: "≈ within 1 cm", add: 9 }, C: { label: "within ½ cm", add: 20 } };
const GSD_MC_TYPES = ["B/A", "B/H", "B/T", "CHECK", "D/N CH", "D/N LS", "F/LATR", "F/LCB", "F/LFB", "F/LFO", "F/LTR", "FLUT", "FOA", "L/H", "M/N", "O/L", "O/L CB", "PICOT", "S/N", "S/N SC", "SNP", "TRIM", "Z/Z"];
const GSD_SEAM_TYPES = ["1th", "2th", "3th", "4th", "5th", "6th"];
const GSD_SKILLS = ["A", "B", "C", "D"];

function gsdElementTMU(el, op) {
  const freq = num(el.freq);
  if (el.type === "mn") return num(el.cm) * freq;
  if (el.type === "sew") {
    const spc = num(op.stitchespercm), rpm = num(op.rpm), len = num(el.cm);
    if (!spc || !rpm) return 0;
    const base = spc / (rpm * 0.0006);
    const ratio = rpm / spc;
    const mult = ratio > 445 ? (Math.pow(4.5 - base, 2) / 100) + 1 : 1;
    const factor = GSD_TENSION[el.tension]?.factor ?? 1;
    const add = GSD_STOP[el.stop]?.add ?? 0;
    return (base * mult * factor * len + 17 + add) * freq;
  }
  const entry = GSD_CODE_LIBRARY[el.code];
  if (!entry) return 0;
  const val = entry.tmu[el.cm];
  return val === undefined ? 0 : val * freq;
}
function gsdElementDesc(el) {
  if (el.type === "mn") return "MANUAL TIME ENTRY";
  if (el.type === "sew") return `SEW ${el.cm || 0} CM`;
  const entry = GSD_CODE_LIBRARY[el.code];
  return entry ? entry.desc : "—";
}
// Rolls a GSD operation's elements up to a SAM: (manual + sew TMU)/2000 + bundle time, x allowance%.
function computeGsdOp(op) {
  let manualTmu = 0, sewTmu = 0, sewLen = 0;
  for (const el of op.elements || []) {
    const tmu = gsdElementTMU(el, op);
    if (el.type === "sew") { sewTmu += tmu; sewLen += num(el.cm) * num(el.freq); }
    else manualTmu += tmu;
  }
  const manualMin = manualTmu / 2000, sewMin = sewTmu / 2000, basicMin = manualMin + sewMin;
  const bundle = num(op.bundle), allowance = num(op.allowance) / 100;
  const smv = (bundle + basicMin) * (1 + allowance);
  return { manualMin, sewMin, basicMin, smv, capacity: smv > 0 ? 60 / smv : null, sewLen };
}
function gsdMacroSmv(macro, findOp) {
  return (macro.items || []).reduce((s, it) => {
    const op = findOp(it.refId);
    return s + (op ? computeGsdOp(op).smv : 0) * num(it.qty || 1);
  }, 0);
}
function gsdFeatureSmv(feature, findOp, findMacro) {
  return (feature.items || []).reduce((s, it) => {
    let each = 0;
    if (it.refType === "macro") { const m = findMacro(it.refId); each = m ? gsdMacroSmv(m, findOp) : 0; }
    else { const o = findOp(it.refId); each = o ? computeGsdOp(o).smv : 0; }
    return s + each * num(it.qty || 1);
  }, 0);
}
function newGsdOperation(n) {
  return {
    id: uid(), code: "OP-" + String(n + 1).padStart(4, "0"), component: "", description: "",
    mctype: GSD_MC_TYPES[14], seamtype: GSD_SEAM_TYPES[0], skill: GSD_SKILLS[0],
    stitchespercm: 4.5, rpm: 4000, bundle: 0, allowance: 15, elements: [],
  };
}
function newGsdElement(type) { return { id: uid(), type: type || "lib", code: "MG2T", tension: "N", stop: "A", cm: 0, freq: 1 }; }
function newGsdMacro(n) { return { id: uid(), code: "MC-" + String(n + 1).padStart(4, "0"), name: "", items: [] }; }
function newGsdFeature(n) { return { id: uid(), code: "FT-" + String(n + 1).padStart(4, "0"), name: "", items: [] }; }

function csvEscape(v) { const s = String(v ?? ""); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
function downloadCsv(filename, columns, rows) {
  const header = columns.map((c) => csvEscape(c[0])).join(",");
  const body = rows.map((r) => columns.map((c) => csvEscape(c[1](r))).join(",")).join("\n");
  const blob = new Blob([header + "\n" + body], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

// Real .xlsx export/import (SheetJS) — requires `npm install xlsx` in the project.
function downloadXlsx(filename, sheetName, columns, rows) {
  const data = rows.map((r) => {
    const obj = {};
    columns.forEach((c) => { obj[c[0]] = c[1](r); });
    return obj;
  });
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName || "Sheet1");
  XLSX.writeFile(wb, filename);
}
// Reads the first sheet of an uploaded .xlsx/.xls file into an array of row objects (header row → keys).
function readXlsxFile(file, onRows, onError) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      onRows(XLSX.utils.sheet_to_json(ws, { defval: "" }));
    } catch (err) {
      (onError || ((m) => alert(m)))("Could not read that Excel file: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);

}
// Full-design Excel report export. The free SheetJS build (`xlsx` package) used above for plain
// data round-tripping cannot write cell colors/borders/fonts — that's a Pro-only feature of that
// library. Excel *does* fully respect inline CSS on an HTML table saved as .xls (no paid library
// needed), so the styled Operation Bulletin report below is built as HTML.
//
// IMPORTANT — multi-sheet correctness: simply putting several <table> elements one after another
// in a single <body> and declaring matching <x:ExcelWorksheet> entries does NOT reliably split
// them into separate tabs — Excel has been seen dumping every table onto the FIRST sheet and
// leaving the rest blank. The only reliable way to get each table on its own real worksheet tab
// is the actual format Excel itself produces for "Web Page, Single File (.mht)": a MIME
// multipart/related message where every sheet is its OWN linked HTML part (Content-Location),
// and a small "root" part declares each <x:ExcelWorksheet><x:WorksheetSource HRef="sheetN.htm"/>
// pointing at that part. That's what buildMhtWorkbook does below.
function escapeHtml(s) {
  return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function buildMhtWorkbook(list) {
  const boundary = `----=_NextPart_${Math.random().toString(36).slice(2)}`;
  const sheetFiles = list.map((sh, i) => `sheet${String(i + 1).padStart(3, "0")}.htm`);
  const worksheetDecls = list.map((sh, i) => `<x:ExcelWorksheet>
<x:Name>${escapeHtml(sh.name)}</x:Name>
<x:WorksheetSource HRef="${sheetFiles[i]}"/>
</x:ExcelWorksheet>`).join("\n");
  const linkTags = sheetFiles.map((f) => `<link rel="File-List" href="${f}">`).join("\n");
  const rootHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="ProgId" content="Excel.Sheet">
${linkTags}
<!--[if gte mso 9]><xml>
<x:ExcelWorkbook>
<x:ExcelWorksheets>
${worksheetDecls}
</x:ExcelWorksheets>
</x:ExcelWorkbook>
</xml><![endif]-->
</head>
<body></body>
</html>`;

  const parts = [
    { location: "root.htm", html: rootHtml },
    ...list.map((sh, i) => ({
      location: sheetFiles[i],
      html: `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body>${sh.html}</body>
</html>`,
    })),
  ];

  const CRLF = "\r\n";
  let body = `MIME-Version: 1.0${CRLF}Content-Type: multipart/related; boundary="${boundary}"${CRLF}${CRLF}`;
  body += `This is a multi-part message using the Excel workbook MIME format.${CRLF}${CRLF}`;
  parts.forEach((p) => {
    body += `--${boundary}${CRLF}`;
    body += `Content-Type: text/html; charset="utf-8"${CRLF}`;
    body += `Content-Location: ${p.location}${CRLF}${CRLF}`;
    body += p.html + CRLF + CRLF;
  });
  body += `--${boundary}--${CRLF}`;
  return body;
}
// ---- Genuine styled .xlsx export (ExcelJS) --------------------------------------------------
// Reads the inline CSS already present on each report's HTML <table> (see obReportStyles above)
// and reproduces it as real Excel formatting — background fill, borders, bold, alignment — plus
// colSpan/rowSpan as merged cells and any <img> as a real embedded picture. This keeps ONE source
// of truth for the report design (the HTML/CSS in buildOperationBulletinReportHtml etc.) instead
// of maintaining a second, separate "Excel version" of every report by hand.
function cssColorToArgb(str) {
  if (!str) return null;
  const s = str.trim();
  let m = s.match(/^#?([0-9a-fA-F]{6})$/);
  if (m) return "FF" + m[1].toUpperCase();
  m = s.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (m) {
    const toHex = (n) => Number(n).toString(16).padStart(2, "0").toUpperCase();
    return "FF" + toHex(m[1]) + toHex(m[2]) + toHex(m[3]);
  }
  return null;
}
function borderSideFromStyle(widthStr, styleStr, colorStr) {
  const px = parseFloat(widthStr) || 0;
  if (!px || !styleStr || styleStr === "none") return undefined;
  const argb = cssColorToArgb(colorStr) || "FF000000";
  return { style: px >= 2 ? "medium" : "thin", color: { argb } };
}
async function htmlTableIntoWorksheet(ws, tableEl, workbookForImages) {
  const rows = Array.from(tableEl.rows);
  const occupied = new Set();
  let maxCol = 0;
  const colTextLen = {};
  for (let r = 0; r < rows.length; r++) {
    const tr = rows[r];
    let c = 0;
    for (const td of Array.from(tr.cells)) {
      while (occupied.has(`${r},${c}`)) c++;
      const rowSpan = td.rowSpan || 1;
      const colSpan = td.colSpan || 1;
      const st = td.style;
      const img = td.querySelector && td.querySelector("img");
      const rawText = (td.textContent || "").trim();
      const upper = (st.textTransform || "").toLowerCase() === "uppercase";
      const cellRef = ws.getCell(r + 1, c + 1);
      if (!img) cellRef.value = upper ? rawText.toUpperCase() : rawText;
      const bg = cssColorToArgb(st.backgroundColor);
      if (bg) cellRef.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
      const bold = st.fontWeight === "bold" || Number(st.fontWeight) >= 600;
      const sizePx = parseFloat(st.fontSize) || 11;
      cellRef.font = { bold, size: Math.max(8, Math.round(sizePx * 0.75)) };
      cellRef.alignment = {
        horizontal: st.textAlign || (img ? "center" : undefined),
        vertical: st.verticalAlign === "middle" ? "middle" : "top",
        wrapText: true,
      };
      const border = {
        top: borderSideFromStyle(st.borderTopWidth, st.borderTopStyle, st.borderTopColor),
        bottom: borderSideFromStyle(st.borderBottomWidth, st.borderBottomStyle, st.borderBottomColor),
        left: borderSideFromStyle(st.borderLeftWidth, st.borderLeftStyle, st.borderLeftColor),
        right: borderSideFromStyle(st.borderRightWidth, st.borderRightStyle, st.borderRightColor),
      };
      if (border.top || border.bottom || border.left || border.right) cellRef.border = border;
      colTextLen[c] = Math.max(colTextLen[c] || 0, rawText.length);

      if (rowSpan > 1 || colSpan > 1) {
        ws.mergeCells(r + 1, c + 1, r + rowSpan, c + colSpan);
        for (let rr = r; rr < r + rowSpan; rr++)
          for (let cc = c; cc < c + colSpan; cc++) {
            occupied.add(`${rr},${cc}`);
            if (rr !== r || cc !== c) {
              const mc = ws.getCell(rr + 1, cc + 1);
              mc.fill = cellRef.fill; mc.border = cellRef.border;
            }
          }
      }
      if (img && img.src && workbookForImages) {
        const extMatch = img.src.match(/^data:image\/(\w+);base64,/);
        const ext = (extMatch ? extMatch[1] : "png").toLowerCase();
        try {
          const imgId = workbookForImages.addImage({ base64: img.src, extension: ext === "jpg" ? "jpeg" : ext });
          ws.addImage(imgId, { tl: { col: c, row: r }, br: { col: c + colSpan, row: r + rowSpan } });
        } catch (e) { /* unsupported image format — leave cell blank rather than fail the export */ }
      }
      maxCol = Math.max(maxCol, c + colSpan);
      c += colSpan;
    }
  }
  for (let c = 0; c < maxCol; c++) {
    ws.getColumn(c + 1).width = Math.min(38, Math.max(9, (colTextLen[c] || 8) + 3));
  }
}
async function downloadHtmlAsExcel(filename, sheets) {
  // Accept either a single HTML table string (old call shape) or an array of { name, html }.
  const list = typeof sheets === "string" ? [{ name: "Sheet1", html: sheets }] : sheets;
  const wb = new ExcelJS.Workbook();
  const usedNames = new Set();
  for (const sh of list) {
    const container = document.createElement("div");
    container.innerHTML = sh.html;
    const table = container.querySelector("table") || container;
    let safeName = (sh.name || "Sheet1").replace(/[\\/?*\[\]:]/g, " ").trim().slice(0, 31) || "Sheet1";
    let n = safeName, i = 2;
    while (usedNames.has(n)) { n = `${safeName.slice(0, 28)} ${i++}`; }
    usedNames.add(n);
    const ws = wb.addWorksheet(n);
    await htmlTableIntoWorksheet(ws, table, wb);
  }
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename.replace(/\.xls$/i, ".xlsx");
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}


// Full-design Operation Bulletin report, styled to match a factory OB sheet (title band, style
// info block, orange column headers, bordered data grid, heading rows, and Section/SMV/Machine
// summary footers) — built as HTML and downloaded via downloadHtmlAsExcel above.
function obReportStyles() {
  return {
    title: "font-family:Calibri,Arial,sans-serif;font-size:20px;font-weight:bold;text-align:center;vertical-align:middle;background:#FFFFFF;border:2px solid #B45F06;padding:10px;",
    bar: "font-family:Calibri,Arial,sans-serif;background:#FCD5B4;font-weight:bold;font-size:13px;border:1px solid #B45F06;padding:5px 8px;",
    infoLabel: "font-family:Calibri,Arial,sans-serif;font-weight:bold;font-size:11px;border:1px solid #D9B48F;padding:4px 8px;background:#FFF6EC;white-space:nowrap;",
    infoValue: "font-family:Calibri,Arial,sans-serif;font-size:11px;border:1px solid #D9B48F;padding:4px 8px;",
    colHeader: "font-family:Calibri,Arial,sans-serif;background:#FCD5B4;font-weight:bold;font-size:11px;text-align:center;border:1px solid #B45F06;padding:6px 4px;",
    cell: "font-family:Calibri,Arial,sans-serif;font-size:11px;border:1px solid #ddd;padding:4px 6px;",
    numCell: "font-family:Calibri,Arial,sans-serif;font-size:11px;border:1px solid #ddd;padding:4px 6px;text-align:center;",
    heading: "font-family:Calibri,Arial,sans-serif;background:#FFF1E0;font-weight:bold;font-size:12px;border:1px solid #ddd;padding:6px 8px;text-transform:uppercase;",
    sectionTitle: "font-family:Calibri,Arial,sans-serif;background:#FCD5B4;font-weight:bold;font-size:12px;border:1px solid #B45F06;padding:5px 8px;",
    totalRow: "font-family:Calibri,Arial,sans-serif;background:#FCE8D6;font-weight:bold;font-size:11px;border:1px solid #B45F06;padding:5px 6px;",
  };
}
// Extra reference photos (style.extraPhotoRefs) beyond the one main Garment Sketch / Photo — laid
// out as its own "ADDITIONAL PHOTOS" band at the bottom of a report table, one <img> per <td> (the
// Excel exporter in htmlTableIntoWorksheet only embeds the FIRST <img> it finds inside a cell, so
// each extra photo needs its own cell rather than being stacked inside the main picture cell).
function buildExtraPhotosBlockHtml(style, s, cols, thumbPx = 90) {
  const extras = (Array.isArray(style.extraPhotoRefs) ? style.extraPhotoRefs : []).filter(Boolean);
  if (!extras.length) return "";
  const rows = [];
  for (let i = 0; i < extras.length; i += cols) {
    const chunk = extras.slice(i, i + cols);
    const cells = chunk
      .map(
        (uri) =>
          `<td style="${s.cell}text-align:center;vertical-align:middle;padding:6px;"><img src="${uri}" alt="Extra Photo" style="max-width:${thumbPx}px;max-height:${thumbPx}px;width:auto;height:auto;object-fit:cover;" /></td>`
      )
      .join("");
    const pad = cols - chunk.length;
    rows.push(`<tr>${cells}${pad > 0 ? `<td colspan="${pad}" style="${s.cell}"></td>` : ""}</tr>`);
  }
  return `
    <tr><td colspan="${cols}" style="height:6px;border:none;"></td></tr>
    <tr><td colspan="${cols}" style="${s.bar}">ADDITIONAL PHOTOS</td></tr>
    ${rows.join("")}`;
}
// Renders a factory-style "LINE BALANCING GRAPH" — a blue Operation Capacity line across every
// operation, plus flat red (UCL) and green (LCL) reference lines — as a self-contained SVG string.
// Used two ways: embedded live as inline <svg> for on-screen Preview/Print, or rasterised to a PNG
// (see svgToPngDataUri) so it can be embedded as a real picture in the Excel export.
function buildLineBalancingGraphSvg(balancing) {
  const rows = (balancing?.rows || []).filter((r) => !r.isHeading);
  if (!rows.length) return { svg: "", width: 0, height: 0 };
  const W = Math.max(760, 110 + rows.length * 75);
  const H = 380;
  const padL = 55, padR = 20, padT = 55, padB = 85;
  const plotW = W - padL - padR;
  const plotH = H - padT - padB;

  const vals = [...rows.map((r) => r.capacityPerHr || 0), balancing.ucl || 0, balancing.lcl || 0].filter((v) => v > 0);
  const rawMax = Math.max(...vals, 1);
  const rawMin = Math.min(...vals, rawMax);
  const pad = Math.max((rawMax - rawMin) * 0.25, rawMax * 0.1, 10);
  const yMax = rawMax + pad;
  const yMin = Math.max(0, rawMin - pad);
  const span = yMax - yMin || 1;
  const yOf = (v) => padT + plotH - ((v - yMin) / span) * plotH;
  const xStep = rows.length > 1 ? plotW / (rows.length - 1) : 0;
  const xOf = (i) => padL + (rows.length > 1 ? i * xStep : plotW / 2);

  const TICKS = 5;
  const gridLines = Array.from({ length: TICKS + 1 }, (_, i) => {
    const v = yMin + (span * i) / TICKS;
    const y = yOf(v);
    return `<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W - padR}" y2="${y.toFixed(1)}" stroke="#e5e0d5" stroke-width="1" />
      <text x="${padL - 6}" y="${(y + 3).toFixed(1)}" font-size="10" text-anchor="end" fill="#6b6455" font-family="Calibri,Arial,sans-serif">${Math.round(v)}</text>`;
  }).join("");

  const linePoints = rows.map((r, i) => `${xOf(i).toFixed(1)},${yOf(r.capacityPerHr || 0).toFixed(1)}`).join(" ");
  const dots = rows.map((r, i) => {
    const x = xOf(i), y = yOf(r.capacityPerHr || 0);
    return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3" fill="#2E75B6" />
      <text x="${x.toFixed(1)}" y="${(y - 7).toFixed(1)}" font-size="9" text-anchor="middle" fill="#2E75B6" font-family="Calibri,Arial,sans-serif">${Math.round(r.capacityPerHr || 0)}</text>`;
  }).join("");

  const xLabels = rows.map((r, i) => {
    const x = xOf(i);
    const words = (r.name || `Op ${r.slNo}`).toUpperCase().split(/\s+/).filter(Boolean);
    const lines = [];
    let cur = "";
    words.forEach((w) => {
      if ((cur ? cur + " " + w : w).length > 11) { if (cur) lines.push(cur); cur = w; }
      else cur = cur ? cur + " " + w : w;
    });
    if (cur) lines.push(cur);
    const shown = lines.slice(0, 2);
    const tspans = shown.map((ln, li) => `<tspan x="${x.toFixed(1)}" dy="${li === 0 ? 0 : 10}">${escapeHtml(ln)}</tspan>`).join("");
    return `<text x="${x.toFixed(1)}" y="${H - padB + 14}" font-size="8.5" text-anchor="middle" fill="#4a4438" font-family="Calibri,Arial,sans-serif">${tspans}</text>`;
  }).join("");

  const uclY = balancing.ucl > 0 ? yOf(balancing.ucl) : null;
  const lclY = balancing.lcl > 0 ? yOf(balancing.lcl) : null;
  const refLines = `
    ${uclY != null ? `<line x1="${padL}" y1="${uclY.toFixed(1)}" x2="${W - padR}" y2="${uclY.toFixed(1)}" stroke="#C0504D" stroke-width="2" stroke-dasharray="6,3" />` : ""}
    ${lclY != null ? `<line x1="${padL}" y1="${lclY.toFixed(1)}" x2="${W - padR}" y2="${lclY.toFixed(1)}" stroke="#4CAF50" stroke-width="2" stroke-dasharray="6,3" />` : ""}`;

  const legend = `
    <g font-family="Calibri,Arial,sans-serif" font-size="10">
      <line x1="${W - 205}" y1="18" x2="${W - 185}" y2="18" stroke="#2E75B6" stroke-width="2" /><text x="${W - 179}" y="21.5" fill="#333">OPERATION CAPACITY</text>
      <line x1="${W - 205}" y1="33" x2="${W - 185}" y2="33" stroke="#C0504D" stroke-width="2" stroke-dasharray="6,3" /><text x="${W - 179}" y="36.5" fill="#333">UCL (${Math.round(balancing.ucl || 0)})</text>
      <line x1="${W - 205}" y1="48" x2="${W - 185}" y2="48" stroke="#4CAF50" stroke-width="2" stroke-dasharray="6,3" /><text x="${W - 179}" y="51.5" fill="#333">LCL (${Math.round(balancing.lcl || 0)})</text>
    </g>`;

  const svg = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg">
    <rect x="0" y="0" width="${W}" height="${H}" fill="#ffffff" />
    <text x="${W / 2}" y="16" font-size="13" font-weight="bold" text-anchor="middle" fill="#375623" font-family="Calibri,Arial,sans-serif">LINE BALANCING GRAPH</text>
    ${gridLines}
    <rect x="${padL}" y="${padT}" width="${plotW}" height="${plotH}" fill="none" stroke="#c9c2b0" />
    ${refLines}
    <polyline points="${linePoints}" fill="none" stroke="#2E75B6" stroke-width="2" />
    ${dots}
    ${xLabels}
    ${legend}
  </svg>`;
  return { svg, width: W, height: H };
}
// Rasterises the SVG above into a PNG data URI via an offscreen <canvas> — only needed for the
// Excel export path, since ExcelJS can embed jpeg/png/gif pictures but not native charts or SVG.
// Async (Image loading is async), so this is only awaited at export time (see handleObExportExcel),
// not during the synchronous HTML building used for on-screen Preview/Print (those embed the live
// SVG directly instead, which needs no conversion).
function svgToPngDataUri(svgMarkup, width, height) {
  return new Promise((resolve) => {
    try {
      const svgBlob = new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" });
      const url = URL.createObjectURL(svgBlob);
      const img = new Image();
      img.onload = () => {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext("2d");
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, width, height);
          ctx.drawImage(img, 0, 0, width, height);
          resolve(canvas.toDataURL("image/png"));
        } catch (e) {
          resolve("");
        } finally {
          URL.revokeObjectURL(url);
        }
      };
      img.onerror = () => { URL.revokeObjectURL(url); resolve(""); };
      img.src = url;
    } catch (e) {
      resolve("");
    }
  });
}
function buildOperationBulletinReportHtml(style, operations, summary, mpBalance, balancing, machineUsage, helperSummary, includeHeadings = true, chartImageUri = "") {
  const s = obReportStyles();
  const td = (v, styleStr) => `<td style="${styleStr}">${escapeHtml(v)}</td>`;
  const COLS = 12; // Sl No, Operation, Section, Machine, SAM, 100% Cap, Plan Cap, OP, HP, Manpower, Eff %, Cap/Hr
  const mpById = Object.fromEntries((mpBalance?.rows || []).map((r) => [r.id, r]));
  const capById = Object.fromEntries((balancing?.rows || []).map((r) => [r.id, r]));
  const totalManpower = (operations || []).filter((o) => !o.isHeading).reduce((a, o) => a + allocatedMpOf(o), 0);

  const photoCellHtml = style.photoRef
    ? `<img src="${style.photoRef}" alt="Garment Sketch" style="max-width:260px;max-height:260px;width:auto;height:auto;object-fit:contain;" />`
    : `<span style="color:#bbb;font-size:11px;">No photo uploaded</span>`;

  const infoRows = `
    <tr>
      ${td("STYLE NO", s.infoLabel)}${td(style.styleNo || "", s.infoValue)}
      ${td("BUYER", s.infoLabel)}${td(style.buyer || "", s.infoValue)}
      ${td("DESCRIPTION", s.infoLabel)}${td(style.obDescription || "", s.infoValue)}
      <td rowspan="4" colspan="6" style="${s.cell}text-align:center;vertical-align:middle;padding:8px;">${photoCellHtml}</td>
    </tr>
    <tr>
      ${td("ORDER QTY", s.infoLabel)}${td(style.orderQty || "", s.infoValue)}
      ${td("TARGET QTY / DAY", s.infoLabel)}${td(style.obTargetQtyPerDay || "", s.infoValue)}
      ${td("PLAN EFF %", s.infoLabel)}${td(style.targetEfficiencyPct ? `${style.targetEfficiencyPct}%` : "", s.infoValue)}
    </tr>
    <tr>
      ${td("SMV (SEW TO QC)", s.infoLabel)}${td(money(summary.sewToQcSmv), s.infoValue)}
      ${td("SMV (SEW TO PACK)", s.infoLabel)}${td(summary.sections.packing.count > 0 ? money(summary.sewToPackSmv) : "", s.infoValue)}
      ${td("TOTAL SMV", s.infoLabel)}${td(money(summary.totalSmv), s.infoValue)}
    </tr>
    <tr>
      ${td("TOTAL OPERATORS", s.infoLabel)}${td(money(helperSummary.totalOp), s.infoValue)}
      ${td("TOTAL HELPERS", s.infoLabel)}${td(money(helperSummary.totalHp), s.infoValue)}
      ${td("TOTAL MANPOWER", s.infoLabel)}${td(money(totalManpower), s.infoValue)}
    </tr>`;

  const colHeaderRow = !includeHeadings ? "" : `
    <tr>
      ${td("SL NO", s.colHeader)}${td("OPERATION", s.colHeader)}${td("SECTION", s.colHeader)}
      ${td("MACHINE", s.colHeader)}${td("SAM", s.colHeader)}${td("100% CAP", s.colHeader)}
      ${td("PLAN CAP", s.colHeader)}${td("OP", s.colHeader)}${td("HP", s.colHeader)}
      ${td("MANPOWER", s.colHeader)}${td("EFF %", s.colHeader)}${td("CAP / HR", s.colHeader)}
    </tr>`;

  const dataRows = (operations || []).map((op) => {
    if (op.isHeading) {
      return `<tr>${td(op.slNo || "", s.numCell)}<td colspan="${COLS - 1}" style="${s.heading}">${escapeHtml(op.name)}</td></tr>`;
    }
    const mp = mpById[op.id];
    const cap = capById[op.id];
    return `<tr>
      ${td(op.slNo || "", s.numCell)}
      ${td(op.name || "", s.cell)}
      ${td((OB_SECTIONS.find((x) => x.key === op.section) || {}).label || op.section || "", s.numCell)}
      ${td(op.machine || "", s.numCell)}
      ${td(op.sam || "", s.numCell)}
      ${td(mp ? money(mp.cap100PerHr) : "—", s.numCell)}
      ${td(mp ? money(mp.planCapPerHr) : "—", s.numCell)}
      ${td(op.opCount || "", s.numCell)}
      ${td(op.hpCount || "", s.numCell)}
      ${td(op.manpower || "", s.numCell)}
      ${td(mp ? `${money(mp.effPct)}%` : "—", s.numCell)}
      ${td(cap ? money(cap.capacityPerHr) : "—", s.numCell)}
    </tr>`;
  }).join("");

  const balancingBlock = balancing.rows.length > 0 ? `
    <tr>${td("BOTTLENECK OPERATION", s.infoLabel)}${td(balancing.bottleneck?.name || "", s.infoValue)}</tr>
    <tr>${td("LCL (PCS/HR)", s.infoLabel)}${td(money(balancing.lcl), s.infoValue)}</tr>
    <tr>${td("UCL (PCS/HR)", s.infoLabel)}${td(money(balancing.ucl), s.infoValue)}</tr>
    <tr>${td("AVG CAPACITY (PCS/HR)", s.infoLabel)}${td(money(balancing.avgCapacity), s.infoValue)}</tr>` : "";

  // A real line graph (Operation Capacity zig-zag + UCL/LCL reference lines), styled like a factory
  // "LINE BALANCING GRAPH". ExcelJS can only embed raster images (jpeg/png/gif), not native charts or
  // inline <svg>, so when the caller has pre-rasterised the graph to a PNG data URI (see
  // svgToPngDataUri, used only at Excel-export time) we embed that as an <img> — same mechanism as
  // the Garment Sketch photo. For on-screen Preview / Print (no Excel round-trip involved) we embed
  // the live <svg> directly, which renders crisper and needs no extra conversion step.
  const graph = buildLineBalancingGraphSvg(balancing);
  const chartMediaHtml = chartImageUri
    ? `<img src="${chartImageUri}" alt="Line Balancing Graph" style="width:100%;max-width:${graph.width}px;height:auto;" />`
    : graph.svg;
  const balancingChartBlock = balancing.rows.length > 0 ? `
    <tr><td colspan="${COLS}" style="${s.sectionTitle}">LINE BALANCING GRAPH (Operation Capacity, pcs/hr — blue = capacity, red = UCL, green = LCL)</td></tr>
    <tr><td colspan="${COLS}" style="${s.cell}text-align:center;padding:6px;">${chartMediaHtml}</td></tr>` : "";

  return `<table style="border-collapse:collapse;">
    <tr><td colspan="${COLS}" style="${s.title}">OPERATION BULLETIN</td></tr>
    <tr><td colspan="6" style="${s.bar}">STYLE DETAILS</td><td colspan="6" style="${s.bar}">GARMENT SKETCH / PICTURE</td></tr>
    ${infoRows}
    <tr><td colspan="${COLS}" style="height:6px;border:none;"></td></tr>
    ${colHeaderRow}
    ${dataRows}
    <tr><td colspan="${COLS}" style="height:10px;border:none;"></td></tr>
    <tr><td colspan="4" style="${s.sectionTitle}">SECTION-WISE SUMMARY</td><td colspan="${COLS - 4}" style="${s.sectionTitle}">MACHINE-WISE USAGE</td></tr>
    ${!includeHeadings ? "" : `<tr>
      ${td("SECTION", s.colHeader)}${td("COUNT", s.colHeader)}${td("SMV", s.colHeader)}${td("MANPOWER", s.colHeader)}
      ${td("MACHINE", s.colHeader)}${td("OP COUNT", s.colHeader)}<td colspan="${COLS - 6}" style="${s.colHeader}">MANPOWER</td>
    </tr>`}
    ${(() => {
      const secArr = OB_SECTIONS.filter((sec) => summary.sections[sec.key].count > 0);
      const maxRows = Math.max(secArr.length, (machineUsage || []).length, 1);
      let rows = "";
      for (let i = 0; i < maxRows; i++) {
        const sec = secArr[i];
        const mach = (machineUsage || [])[i];
        rows += `<tr>
          ${sec ? td(sec.label, s.cell) + td(summary.sections[sec.key].count, s.numCell) + td(money(summary.sections[sec.key].smv), s.numCell) + td(money(summary.sections[sec.key].manpower), s.numCell) : td("", s.cell) + td("", s.numCell) + td("", s.numCell) + td("", s.numCell)}
          ${mach ? td(mach.machine, s.cell) + td(mach.opCount, s.numCell) + `<td colspan="${COLS - 6}" style="${s.numCell}">${money(mach.manpower)}</td>` : td("", s.cell) + td("", s.numCell) + `<td colspan="${COLS - 6}" style="${s.numCell}"></td>`}
        </tr>`;
      }
      return rows;
    })()}
    <tr>
      ${td("TOTAL", s.totalRow)}${td(Object.values(summary.sections).reduce((a, x) => a + x.count, 0), s.totalRow)}
      ${td(money(summary.totalSmv), s.totalRow)}${td(money(totalManpower), s.totalRow)}
      <td colspan="${COLS - 4}" style="${s.totalRow}"></td>
    </tr>
    <tr><td colspan="${COLS}" style="height:10px;border:none;"></td></tr>
    <tr><td colspan="${COLS}" style="${s.sectionTitle}">LINE BALANCING</td></tr>
    ${balancingBlock}
    <tr><td colspan="${COLS}" style="height:10px;border:none;"></td></tr>
    ${balancingChartBlock}
    ${buildExtraPhotosBlockHtml(style, s, COLS, 100)}
  </table>`;
}
// Ramp-up Plan sheet — same day-wise plan as the in-app Ramp-up Plan section (manual entries for
// early days, auto-filled at Line Target for the rest), styled to match the Operation Bulletin sheet.
function buildRampUpReportHtml(style, rampUpPlan, rampUpExtras = {}, includeHeadings = true) {
  const s = obReportStyles();
  const td = (v, styleStr) => `<td style="${styleStr}">${escapeHtml(v)}</td>`;
  const COLS = 6; // Day, Qty, Type, Eff % (day), Cumulative, Chart
  const { actualMp100 = 0, avgEfficiencyPct = 0, noOfLines = 0, perLineTarget = 0, totalLineTarget = 0, totalMp = 0, sewingSam = 0, shiftMinutes = SHIFT_MINUTES } = rampUpExtras;
  // Per-day production efficiency — day capacity @ 100% = (Total Manpower × Working Minutes) ÷ SAM;
  // that day's Eff% = day qty ÷ that capacity × 100. Same formula as the in-app Ramp-up Plan list.
  const dayEffPct = (qty) => {
    const capacity100 = totalMp > 0 && sewingSam > 0 ? (totalMp * shiftMinutes) / sewingSam : 0;
    return capacity100 > 0 ? (num(qty) / capacity100) * 100 : 0;
  };
  let cumulative = 0;
  const maxQty = Math.max(1, ...rampUpPlan.rows.map((r) => r.qty));
  const dayRows = rampUpPlan.rows.map((r) => {
    cumulative += r.qty;
    const pct = Math.max(2, Math.round((r.qty / maxQty) * 100));
    const barColor = r.manual ? "#F79646" : "#4F81BD";
    const eff = dayEffPct(r.qty);
    return `<tr>
      ${td(`Day ${r.day}`, s.cell)}
      ${td(money(r.qty), s.numCell)}
      ${td(r.manual ? "Manual" : "Auto", s.numCell)}
      ${td(eff ? `${money(eff)}%` : "", s.numCell)}
      ${td(money(cumulative), s.numCell)}
      <td style="${s.cell}padding:3px 6px;"><div style="background:${barColor};height:14px;width:${pct}%;min-width:4px;"></div></td>
    </tr>`;
  }).join("");
  return `<table style="border-collapse:collapse;">
    <tr><td colspan="${COLS}" style="${s.title}">RAMP-UP PLAN</td></tr>
    <tr><td colspan="${COLS}" style="${s.bar}">STYLE DETAILS</td></tr>
    <tr>
      ${td("STYLE NO", s.infoLabel)}${td(style.styleNo || "", s.infoValue)}
      ${td("ORDER QTY", s.infoLabel)}${td(style.orderQty || "", s.infoValue)}
      ${td("TOTAL LINE TARGET / DAY", s.infoLabel)}
    </tr>
    <tr>
      ${td("QTY ENTERED (MANUAL)", s.infoLabel)}${td(money(rampUpPlan.manualSum), s.infoValue)}
      ${td("REMAINING QTY", s.infoLabel)}${td(money(rampUpPlan.remaining), s.infoValue)}
      ${td(totalLineTarget ? money(totalLineTarget) : "", s.infoValue)}
    </tr>
    <tr>
      ${td("DAYS NEEDED (EXACT)", s.infoLabel)}${td(rampUpPlan.autoDaysExact.toFixed(1), s.infoValue)}
      ${td("LEAD TIME — TOTAL DAYS", s.infoLabel)}${td(rampUpPlan.totalDaysExact.toFixed(1), s.infoValue)}
      ${td("", s.infoValue)}
    </tr>
    <tr><td colspan="${COLS}" style="${s.bar}">MANPOWER &amp; LINE PLANNING</td></tr>
    <tr>
      ${td("TOTAL MANPOWER (CUT+SEW+QC+PACK)", s.infoLabel)}${td(money(actualMp100), s.infoValue)}
      ${td("AVG EFFICIENCY %", s.infoLabel)}${td(avgEfficiencyPct ? `${money(avgEfficiencyPct)}%` : "", s.infoValue)}
    </tr>
    <tr>
      ${td("NO. OF LINES", s.infoLabel)}${td(noOfLines || "", s.infoValue)}
      ${td("TARGET QTY / DAY / LINE", s.infoLabel)}${td(perLineTarget ? money(perLineTarget) : "", s.infoValue)}
    </tr>
    <tr><td colspan="${COLS}" style="height:6px;border:none;"></td></tr>
    ${!includeHeadings ? "" : `<tr>${td("DAY", s.colHeader)}${td("QTY", s.colHeader)}${td("TYPE", s.colHeader)}${td("EFF %", s.colHeader)}${td("CUMULATIVE QTY", s.colHeader)}${td("CHART (orange = manual, blue = auto)", s.colHeader)}</tr>`}
    ${dayRows || `<tr><td colspan="${COLS}" style="${s.cell}text-align:center;">No ramp-up days entered yet.</td></tr>`}
  </table>`;
}
// Product Costing Details sheet — mirrors a factory's Product Costing card (Buyer/Style/Description
// block, Total SMV/Order Eff/Lead Time/Planned Line/Compiled By/Dates block, and a Plan Target /
// Target 100% / Line Balance footer row) so the exported workbook has a costing summary page too.
function buildProductCostingReportHtml(style, summary, balancing, rampUpPlan, achievableQtyAtActualMp, includeHeadings = true) {
  const s = obReportStyles();
  const td = (v, styleStr) => `<td style="${styleStr}">${escapeHtml(v)}</td>`;
  const COLS = 6;
  const orderEffPct = style.targetEfficiencyPct ? `${style.targetEfficiencyPct}%` : "";
  const leadTime = style.obLeadTime || (rampUpPlan.totalDaysExact ? rampUpPlan.totalDaysExact.toFixed(1) : "");
  const target100 = Math.round(achievableQtyAtActualMp || 0);
  const planTarget = num(style.obTargetQtyPerDay);
  const planTargetPct = target100 > 0 ? (planTarget / target100) * 100 : 0;
  const lineBalancePct = balancing.ucl > 0 ? (balancing.lcl / balancing.ucl) * 100 : 0;
  return `<table style="border-collapse:collapse;">
    <tr><td colspan="${COLS}" style="${s.title}">PRODUCT COSTING DETAILS</td></tr>
    ${!includeHeadings ? "" : `<tr>
      <td colspan="2" style="${s.bar}">PRODUCT DETAILS</td>
      <td colspan="2" style="${s.bar}">RESULTS</td>
      <td colspan="2" style="${s.bar}">GARMENT SKETCH / PICTURE</td>
    </tr>`}
    <tr>
      ${td("BUYER", s.infoLabel)}${td(style.buyer || "", s.infoValue)}
      ${td("TOTAL SMV (SEW TO PACK)", s.infoLabel)}${td(money(summary.sewToPackSmv), s.infoValue)}
      <td rowspan="6" colspan="2" style="${s.cell}text-align:center;vertical-align:middle;padding:8px;">${style.photoRef ? `<img src="${style.photoRef}" alt="Garment Sketch" style="max-width:170px;max-height:170px;width:auto;height:auto;object-fit:contain;" />` : `<span style="color:#bbb;font-size:11px;">No photo uploaded</span>`}</td>
    </tr>
    <tr>
      ${td("STYLE NO", s.infoLabel)}${td(style.styleNo || "", s.infoValue)}
      ${td("ORDER EFF %", s.infoLabel)}${td(orderEffPct, s.infoValue)}
    </tr>
    <tr>
      ${td("DESCRIPTION", s.infoLabel)}${td(style.obDescription || "", s.infoValue)}
      ${td("LEAD TIME (DAYS)", s.infoLabel)}${td(leadTime, s.infoValue)}
    </tr>
    <tr>
      ${td("ORDER QTY", s.infoLabel)}${td(style.orderQty || "", s.infoValue)}
      ${td("PLANNED LINE", s.infoLabel)}${td(style.obPlannedLines || "", s.infoValue)}
    </tr>
    <tr>
      ${td("FACTORY", s.infoLabel)}${td(style.obFactory || "", s.infoValue)}
      ${td("COMPILED BY", s.infoLabel)}${td(style.obCompiledBy || "", s.infoValue)}
    </tr>
    <tr>
      ${td("COUNTRY", s.infoLabel)}${td(style.obCountry || "", s.infoValue)}
      ${td("COMPILED DATE", s.infoLabel)}${td(style.obCompiledDate || "", s.infoValue)}
    </tr>
    <tr>
      ${td("", s.infoValue)}${td("", s.infoValue)}
      ${td("LAST CHANGE DATE", s.infoLabel)}${td(style.obLastChangeDate || "", s.infoValue)}
      ${td("", s.infoValue)}${td("", s.infoValue)}
    </tr>
    <tr><td colspan="${COLS}" style="height:6px;border:none;"></td></tr>
    ${!includeHeadings ? "" : `<tr>${td("PLAN TARGET", s.colHeader)}${td("", s.colHeader)}${td("", s.colHeader)}${td("TARGET 100%", s.colHeader)}${td("LINE BALANCE", s.colHeader)}${td("", s.colHeader)}</tr>`}
    <tr>
      ${td(money(planTarget), s.numCell)}${td(`${money(planTargetPct)}%`, s.numCell)}${td("", s.numCell)}
      ${td(money(target100), s.numCell)}${td(`${money(lineBalancePct)}%`, s.numCell)}${td("", s.numCell)}
    </tr>
    ${buildExtraPhotosBlockHtml(style, s, COLS, 90)}
  </table>`;
}
function Stat({ label, value, accent, sub, big }) {
  return (
    <div className="flex flex-col gap-0.5">
      <span className="text-[10px] uppercase tracking-wider text-stone-400 font-semibold">{label}</span>
      <span className={`${big ? "text-2xl" : "text-base"} font-bold tabular-nums leading-tight ${accent || "text-stone-800"}`}>{value}</span>
      {sub && <span className="text-[11px] text-stone-400">{sub}</span>}
    </div>
  );
}
function Field({ label, value, onChange, suffix, type = "number", placeholder }) {
  return (
    <label className="flex flex-col gap-1">
      <span className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">{label}</span>
      <div className="flex items-center rounded-xl border border-stone-200 bg-white transition-colors focus-within:border-amber-500 focus-within:ring-2 focus-within:ring-amber-500/20 hover:border-stone-300">
        <input
          type={type}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          className="w-full px-3 py-2.5 text-sm rounded-xl outline-none bg-transparent tabular-nums placeholder:text-stone-300"
        />
        {suffix && <span className="pr-3 text-[11px] text-stone-400 whitespace-nowrap">{suffix}</span>}
      </div>
    </label>
  );
}
function Section({ title, icon: Icon, children }) {
  const [visible, setVisible] = useState(true);
  return (
    <div>
      <h4 className="flex items-center justify-between gap-1.5 text-xs font-bold text-stone-500 uppercase tracking-wider mb-2.5">
        <span className="flex items-center gap-1.5 min-w-0">
          {Icon && <Icon size={13} className="text-amber-600 shrink-0" />} <span className="truncate">{title}</span>
        </span>
        <button
          type="button"
          onClick={() => setVisible((v) => !v)}
          title={visible ? "Hide this section" : "Show this section"}
          className="shrink-0 text-stone-400 active:text-amber-600 p-0.5"
        >
          {visible ? <Eye size={14} /> : <EyeOff size={14} />}
        </button>
      </h4>
      <div className={visible ? "" : "hidden"}>{children}</div>
    </div>
  );
}
function EmptyState({ icon: Icon, title, body, actionLabel, onAction }) {
  return (
    <div className="flex flex-col items-center text-center gap-3 rounded-2xl border-2 border-dashed border-stone-300 bg-stone-50/60 py-16 px-6">
      <div className="flex items-center justify-center w-12 h-12 rounded-full bg-white border border-stone-200">
        <Icon size={22} className="text-stone-300" />
      </div>
      <div>
        <p className="font-bold text-stone-700 text-sm">{title}</p>
        <p className="text-xs text-stone-400 max-w-xs mt-1">{body}</p>
      </div>
      {actionLabel && (
        <button onClick={onAction} className="mt-1 flex items-center gap-1.5 rounded-xl bg-amber-500 text-stone-900 px-4 py-2.5 text-sm font-bold shadow-sm shadow-amber-500/30 active:bg-amber-400 transition-colors">
          <Plus size={15} /> {actionLabel}
        </button>
      )}
    </div>
  );
}

// Reusable Photo Reference upload + preview control — used on New Style creation and on Add
// Group (Sub Entries), so a reference photo can be attached and previewed right where it's needed.
function PhotoRefUpload({ value, onChange, label = "Photo Reference" }) {
  const inputRef = useRef(null);
  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onChange(reader.result);
    reader.readAsDataURL(file);
    e.target.value = "";
  };
  return (
    <div className="space-y-1.5">
      <p className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">{label}</p>
      <input ref={inputRef} type="file" accept="image/*" onChange={handleFile} className="hidden" />
      {value ? (
        <div className="flex items-center gap-2.5">
          <div className="relative w-20 h-20 rounded-xl overflow-hidden border border-stone-200 shrink-0">
            <img src={value} alt="Reference preview" className="w-full h-full object-cover" />
            <button
              type="button"
              onClick={() => onChange("")}
              className="absolute top-1 right-1 rounded-full bg-black/60 text-white p-0.5 active:bg-black/80"
              title="Remove photo"
            >
              <X size={12} />
            </button>
          </div>
          <button
            type="button"
            onClick={() => inputRef.current?.click()}
            className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3 py-2 text-xs font-bold active:bg-stone-50"
          >
            <Upload size={13} /> Change
          </button>
        </div>
      ) : (
        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          className="flex items-center gap-1.5 rounded-xl border-2 border-dashed border-stone-300 text-stone-500 px-3 py-2 text-xs font-bold active:bg-stone-50"
        >
          <Upload size={14} /> Upload Photo Reference
        </button>
      )}
    </div>
  );
}

// Multi-photo upload + preview grid — used alongside PhotoRefUpload so a style can carry extra
// reference photos (besides the one main Garment Sketch / Photo) that also flow into the
// Product Costing and Operation Bulletin Excel sheets' GARMENT SKETCH / PICTURE block.
function PhotoGalleryUpload({ value, onChange, label = "Extra Photos" }) {
  const inputRef = useRef(null);
  const photos = Array.isArray(value) ? value : [];
  const handleFiles = (e) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    Promise.all(
      files.map(
        (file) =>
          new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.readAsDataURL(file);
          })
      )
    ).then((dataUris) => onChange([...photos, ...dataUris]));
    e.target.value = "";
  };
  const removeAt = (idx) => onChange(photos.filter((_, i) => i !== idx));
  return (
    <div className="space-y-1.5">
      <p className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">{label}</p>
      <input ref={inputRef} type="file" accept="image/*" multiple onChange={handleFiles} className="hidden" />
      <div className="flex flex-wrap items-center gap-2">
        {photos.map((uri, idx) => (
          <div key={idx} className="relative w-16 h-16 rounded-xl overflow-hidden border border-stone-200 shrink-0">
            <img src={uri} alt={`Extra photo ${idx + 1}`} className="w-full h-full object-cover" />
            <button
              type="button"
              onClick={() => removeAt(idx)}
              className="absolute top-0.5 right-0.5 rounded-full bg-black/60 text-white p-0.5 active:bg-black/80"
              title="Remove photo"
            >
              <X size={11} />
            </button>
          </div>
        ))}
        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          className="flex items-center justify-center gap-1 w-16 h-16 rounded-xl border-2 border-dashed border-stone-300 text-stone-500 text-[10px] font-bold active:bg-stone-50 shrink-0"
        >
          <Upload size={13} /> Add
        </button>
      </div>
    </div>
  );
}

// OB Template Library — saved Sl No/Operation/Section/Machine/SAM/Manpower sets keyed by Style
// Description (e.g. "C Neck Long Sleeve"), so a similar new style can load a ready-made bulletin
// instead of building one from scratch every time.
const OB_TEMPLATE_KEY = "sgps-ob-templates";
function loadObTemplatesRaw() {
  try {
    const raw = localStorage.getItem(OB_TEMPLATE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return [];
}
function useObTemplateLibrary() {
  const [templates, setTemplatesState] = useState(loadObTemplatesRaw);
  const save = (next) => {
    setTemplatesState(next);
    try { localStorage.setItem(OB_TEMPLATE_KEY, JSON.stringify(next)); } catch (e) {}
  };
  return { templates, save };
}

// OB Master Template file — the actual factory Excel file (e.g. "MB OB SHEET.xlsx") uploaded once
// and kept on this device, so the whole team knows which format the Operation Bulletin should
// follow. Stored as base64 in localStorage (separate from the per-style OB Template Library above).
const OB_MASTER_TEMPLATE_KEY = "sgps-ob-master-template";
function loadObMasterTemplateRaw() {
  try {
    const raw = localStorage.getItem(OB_MASTER_TEMPLATE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return null;
}
function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}
function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}
function useObMasterTemplate() {
  const [masterTemplate, setMasterTemplateState] = useState(loadObMasterTemplateRaw);
  const save = (next) => {
    setMasterTemplateState(next);
    try {
      if (next) localStorage.setItem(OB_MASTER_TEMPLATE_KEY, JSON.stringify(next));
      else localStorage.removeItem(OB_MASTER_TEMPLATE_KEY);
    } catch (e) {}
  };
  return { masterTemplate, save };
}

const GSD_LIB_KEY = "sgps-gsd-library";
function loadGsdLibraryRaw() {
  try {
    const raw = localStorage.getItem(GSD_LIB_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return { operations: [], macros: [], features: [] };
}
function useGsdLibrary() {
  const [lib, setLib] = useState(loadGsdLibraryRaw);
  const save = (next) => {
    setLib(next);
    try { localStorage.setItem(GSD_LIB_KEY, JSON.stringify(next)); } catch (e) {}
  };
  const findOp = (id) => lib.operations.find((o) => o.id === id);
  const findMacro = (id) => lib.macros.find((m) => m.id === id);
  return { lib, save, findOp, findMacro };
}

// OB Sub Entries Library — reusable operations (Name/Section/Machine/SAM/Manpower), each optionally
// linked to a Heading. Saved once here, they can be inserted into any style straight into that
// heading's group instead of typing the same lines out every time.
const OB_SUBENTRY_KEY = "sgps-ob-subentries";
function loadObSubEntriesRaw() {
  try {
    const raw = localStorage.getItem(OB_SUBENTRY_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return [];
}
function useObSubEntryLibrary() {
  const [entries, setEntriesState] = useState(loadObSubEntriesRaw);
  const save = (next) => {
    setEntriesState(next);
    try { localStorage.setItem(OB_SUBENTRY_KEY, JSON.stringify(next)); } catch (e) {}
  };
  const addEntry = (entry) => save([...entries, { ...entry, id: uid() }]);
  const addEntries = (list) => save([...entries, ...list.map((entry) => ({ ...entry, id: uid() }))]);
  const updateEntry = (id, patch) => save(entries.map((e) => (e.id === id ? { ...e, ...patch } : e)));
  const deleteEntry = (id) => save(entries.filter((e) => e.id !== id));
  const deleteEntries = (ids) => save(entries.filter((e) => !ids.includes(e.id)));
  return { entries, addEntry, addEntries, updateEntry, deleteEntry, deleteEntries };
}
// it shows up as a quick-pick suggestion (datalist) next time, instead of retyping it for every style.
const OB_HEADING_KEY = "sgps-ob-headings";
function loadObHeadingsRaw() {
  try {
    const raw = localStorage.getItem(OB_HEADING_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return [];
}
function useObHeadingLibrary() {
  const [headings, setHeadingsState] = useState(loadObHeadingsRaw);
  const addHeading = (name) => {
    const clean = (name || "").trim();
    if (!clean) return;
    setHeadingsState((prev) => {
      if (prev.some((h) => h.toLowerCase() === clean.toLowerCase())) return prev;
      const next = [...prev, clean];
      try { localStorage.setItem(OB_HEADING_KEY, JSON.stringify(next)); } catch (e) {}
      return next;
    });
  };
  return { headings, addHeading };
}

// Photo reference per Add-Group heading — one reference image saved against a heading name (e.g.
// THUMB HOLE) so the whole group's rows have a picture to work from when re-used on a new style.
const OB_GROUP_PHOTO_KEY = "sgps-ob-group-photos";
function loadObGroupPhotosRaw() {
  try {
    const raw = localStorage.getItem(OB_GROUP_PHOTO_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return {};
}
function useObGroupPhotoLibrary() {
  const [photos, setPhotosState] = useState(loadObGroupPhotosRaw);
  const save = (next) => {
    setPhotosState(next);
    try { localStorage.setItem(OB_GROUP_PHOTO_KEY, JSON.stringify(next)); } catch (e) {}
  };
  const setPhoto = (headingName, dataUri) => {
    const key = (headingName || "").trim().toLowerCase();
    if (!key) return;
    const next = { ...photos };
    if (dataUri) next[key] = dataUri; else delete next[key];
    save(next);
  };
  const getPhoto = (headingName) => photos[(headingName || "").trim().toLowerCase()] || "";
  return { photos, setPhoto, getPhoto };
}

// ---------------- Saved OB — Buyer-wise library, with Print / Share ----------------
// Groups every style that already has OB operations saved, by Buyer, so a saved Operation
// Bulletin can be found and handed out (printed / shared) without hunting through the Style switcher.
function SavedObScreen({ styles, onEditStyle }) {
  const [openBuyer, setOpenBuyer] = useState("");
  const [selectedId, setSelectedId] = useState("");
  const [shareMsg, setShareMsg] = useState("");
  const [search, setSearch] = useState("");

  const savedStyles = useMemo(
    () => (styles || []).filter((s) => (s.obOperations || []).some((op) => !op.isHeading)),
    [styles]
  );
  // Buyer-wise search: matches on Buyer name, Style No, or Description, case-insensitive. While
  // searching, groups are auto-expanded so matches aren't hidden behind a collapsed buyer.
  const searchedStyles = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return savedStyles;
    return savedStyles.filter((s) => {
      const buyer = (s.buyer || "").toLowerCase();
      const styleNo = (s.styleNo || "").toLowerCase();
      const desc = (s.obDescription || "").toLowerCase();
      return buyer.includes(q) || styleNo.includes(q) || desc.includes(q);
    });
  }, [savedStyles, search]);
  const byBuyer = useMemo(() => {
    const map = {};
    searchedStyles.forEach((s) => {
      const b = (s.buyer || "").trim() || "No Buyer";
      if (!map[b]) map[b] = [];
      map[b].push(s);
    });
    return Object.entries(map).sort((a, b) => a[0].localeCompare(b[0]));
  }, [searchedStyles]);

  const selected = savedStyles.find((s) => s.id === selectedId) || null;
  const summary = useMemo(() => computeOBSummary(selected?.obOperations), [selected]);
  const ops = selected?.obOperations || [];

  const shareText = () => {
    if (!selected) return "";
    return [
      `Operation Bulletin — ${selected.styleNo || "Untitled"}`,
      selected.buyer ? `Buyer: ${selected.buyer}` : "",
      selected.obDescription ? `Description: ${selected.obDescription}` : "",
      `Order Qty: ${selected.orderQty || "—"} pcs`,
      selected.obTargetQtyPerDay ? `Target Qty/Day: ${selected.obTargetQtyPerDay} pcs` : "",
      `SMV (Sew to QC): ${money(summary.sewToQcSmv)} · SMV (Sew to Pack): ${money(summary.sewToPackSmv)} · Total SMV: ${money(summary.totalSmv)}`,
      `Operations: ${ops.filter((o) => !o.isHeading).length}`,
    ].filter(Boolean).join("\n");
  };

  const handlePrint = () => { if (selected) window.print(); };

  const handleShare = async () => {
    if (!selected) return;
    const text = shareText();
    try {
      if (navigator.share) {
        await navigator.share({ title: `OB — ${selected.styleNo || "Untitled"}`, text });
      } else {
        await navigator.clipboard.writeText(text);
        setShareMsg("Copied to clipboard");
        setTimeout(() => setShareMsg(""), 2500);
      }
    } catch (e) {
      // share sheet cancelled — nothing to do
    }
  };

  if (savedStyles.length === 0) {
    return (
      <EmptyState
        icon={ClipboardList}
        title="No saved OB yet"
        body="Fill operations in Operation Bulletin and hit Save — it'll show up here, grouped by buyer, ready to print or share."
      />
    );
  }

  return (
    <div className="space-y-3">
      {/* Print-only CSS: hides everything except the #ob-print-area card when window.print() runs */}
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #ob-print-area, #ob-print-area * { visibility: visible; }
          #ob-print-area { position: absolute; left: 0; top: 0; width: 100%; padding: 16px; }
        }
      `}</style>

      <div className="rounded-2xl border border-stone-200 bg-white p-3">
        <p className="text-[11px] font-bold text-stone-500 uppercase tracking-wide mb-2">Saved OB — Buyer-wise</p>
        <div className="flex items-center rounded-xl border border-stone-200 bg-white mb-2.5 focus-within:border-amber-500">
          <Search size={14} className="ml-3 text-stone-400 shrink-0" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search by Buyer, Style No, or Description..."
            className="w-full px-2 py-2.5 text-xs outline-none bg-transparent placeholder:text-stone-300"
          />
          {search && (
            <button onClick={() => setSearch("")} className="mr-3 text-stone-300 active:text-red-600 shrink-0"><X size={14} /></button>
          )}
        </div>
        {byBuyer.length === 0 ? (
          <p className="text-xs text-stone-400 py-2 text-center">No saved OB matches "{search}".</p>
        ) : (
        <div className="space-y-1.5">
          {byBuyer.map(([buyer, list]) => (
            <div key={buyer} className="rounded-xl border border-stone-200 overflow-hidden">
              <button
                onClick={() => setOpenBuyer((b) => (b === buyer ? "" : buyer))}
                className="w-full flex items-center justify-between px-3 py-2.5 bg-stone-50 text-left"
              >
                <span className="text-xs font-extrabold text-stone-700">{buyer}</span>
                <span className="flex items-center gap-2">
                  <span className="text-[10px] font-bold text-stone-400">{list.length} style{list.length > 1 ? "s" : ""}</span>
                  <ChevronDown size={14} className={`text-stone-400 transition-transform ${(openBuyer === buyer || (search.trim() && byBuyer.length <= 3)) ? "" : "-rotate-90"}`} />
                </span>
              </button>
              {(openBuyer === buyer || (search.trim() && byBuyer.length <= 3)) && (
                <div className="divide-y divide-stone-100">
                  {list.map((s) => (
                    <div
                      key={s.id}
                      className={`w-full flex items-center justify-between px-3 py-2.5 text-left gap-2 ${selectedId === s.id ? "bg-amber-50" : "bg-white"}`}
                    >
                      <button onClick={() => setSelectedId(s.id)} className="min-w-0 flex-1 text-left">
                        <span className="block text-xs font-bold text-stone-800 truncate">{s.styleNo || "Untitled"}</span>
                        <span className="block text-[10px] text-stone-400">{(s.obOperations || []).filter((o) => !o.isHeading).length} operations · {s.orderQty || 0} pcs</span>
                      </button>
                      {selectedId === s.id && <Check size={14} className="text-amber-600 shrink-0" />}
                      {onEditStyle && (
                        <button
                          onClick={() => onEditStyle(s.id)}
                          className="flex items-center gap-1 rounded-lg border border-amber-300 text-amber-700 px-2 py-1 text-[10px] font-bold shrink-0 active:bg-amber-100"
                        >
                          <Pencil size={11} /> Edit
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
        )}
      </div>

      {selected && (
        <div className="rounded-2xl border border-stone-200 bg-white p-4">
          <div className="flex flex-wrap items-center gap-2 mb-3">
            {onEditStyle && (
              <button onClick={() => onEditStyle(selected.id)} className="flex items-center gap-1.5 rounded-xl bg-amber-500 text-stone-900 px-3.5 py-2 text-xs font-bold active:bg-amber-400">
                <Pencil size={14} /> Edit
              </button>
            )}
            <button onClick={handlePrint} className="flex items-center gap-1.5 rounded-xl bg-stone-800 text-white px-3.5 py-2 text-xs font-bold active:bg-stone-900">
              <Printer size={14} /> Print
            </button>
            <button onClick={handleShare} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold">
              <Share2 size={14} /> Share
            </button>
            {shareMsg && <span className="flex items-center gap-1 text-xs font-semibold text-emerald-600"><Check size={13} /> {shareMsg}</span>}
          </div>

          <div id="ob-print-area">
            <h3 className="text-sm font-extrabold text-stone-800">{selected.styleNo || "Untitled"}</h3>
            <p className="text-xs text-stone-500 mb-2">
              {selected.buyer ? `Buyer: ${selected.buyer}` : ""}{selected.obDescription ? ` · ${selected.obDescription}` : ""}
            </p>
            <div className="grid grid-cols-3 gap-2.5 mb-3">
              <Stat label="Order Qty" value={selected.orderQty || "—"} />
              <Stat label="Target Qty/Day" value={selected.obTargetQtyPerDay || "—"} />
              <Stat label="Total SMV" value={money(summary.totalSmv)} accent="text-amber-600" />
            </div>
            <table className="w-full text-[11px] border-collapse">
              <thead>
                <tr className="bg-stone-100 text-stone-500 text-left">
                  <th className="px-2 py-1.5 border border-stone-200">SL</th>
                  <th className="px-2 py-1.5 border border-stone-200">Operation</th>
                  <th className="px-2 py-1.5 border border-stone-200">Section</th>
                  <th className="px-2 py-1.5 border border-stone-200">Machine</th>
                  <th className="px-2 py-1.5 border border-stone-200">SAM</th>
                  <th className="px-2 py-1.5 border border-stone-200">MP</th>
                </tr>
              </thead>
              <tbody>
                {ops.map((op) =>
                  op.isHeading ? (
                    <tr key={op.id}>
                      <td colSpan={6} className="px-2 py-1.5 border border-stone-200 bg-amber-50 font-bold text-amber-700">{op.name}</td>
                    </tr>
                  ) : (
                    <tr key={op.id}>
                      <td className="px-2 py-1 border border-stone-200 tabular-nums">{op.slNo}</td>
                      <td className="px-2 py-1 border border-stone-200">{op.name}</td>
                      <td className="px-2 py-1 border border-stone-200 capitalize">{op.section}</td>
                      <td className="px-2 py-1 border border-stone-200">{op.machine}</td>
                      <td className="px-2 py-1 border border-stone-200 tabular-nums">{money(op.sam)}</td>
                      <td className="px-2 py-1 border border-stone-200 tabular-nums">{op.manpower}</td>
                    </tr>
                  )
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ---------------- Operation Bulletin & Line Balancing ----------------
function OperationBulletinScreen({ styles, onUpdateStyle, onCreateStyle, obTab, setObTab }) {
  const [styleId, setStyleId] = useState(styles[0]?.id || "");
  const [savedMsg, setSavedMsg] = useState("");
  const [newStyleNo, setNewStyleNo] = useState("");
  const [newBuyer, setNewBuyer] = useState("");
  const [newOrderQty, setNewOrderQty] = useState("");
  const [newPhotoRef, setNewPhotoRef] = useState("");
  const [bulkHeading, setBulkHeading] = useState("");
  const [bulkPhotoRef, setBulkPhotoRef] = useState("");
  const blankBulkRow = () => ({ id: uid(), name: "", section: "sewing", machine: "", sam: "", opCount: "", hpCount: "" });
  const [bulkRows, setBulkRows] = useState([blankBulkRow()]);
  useEffect(() => { if (!styleId && styles[0]) setStyleId(styles[0].id); }, [styles]);
  const handleCreateStyle = () => {
    const id = onCreateStyle({ styleNo: newStyleNo, buyer: newBuyer, orderQty: newOrderQty, photoRef: newPhotoRef });
    setStyleId(id);
    setObTab("target");
    setNewStyleNo("");
    setNewBuyer("");
    setNewOrderQty("");
    setNewPhotoRef("");
  };

  const style = styles.find((s) => s.id === styleId);
  const operations = style?.obOperations || [];
  const summary = useMemo(() => computeOBSummary(operations), [operations]);
  const balancing = useMemo(() => computeLineBalancing(operations), [operations]);

  const setOperations = (ops) => onUpdateStyle({ ...style, obOperations: ops });
  // Sl No is fully automatic now — always 1,2,3... across the non-heading rows in on-screen order,
  // headings blank. Runs after every insert/delete/reorder so numbers stay sequential even when a
  // row is added in the middle. The equality checks below make this a no-op once numbers already
  // match, so it doesn't loop.
  const operationsNeedRenumber = (list) => {
    let n = 0;
    for (const o of list) {
      if (o.isHeading) { if (o.slNo !== "") return true; continue; }
      n += 1;
      if (o.slNo !== String(n)) return true;
    }
    return false;
  };
  const renumberOperations = (list) => {
    let n = 0;
    return list.map((o) => {
      if (o.isHeading) return o.slNo === "" ? o : { ...o, slNo: "" };
      n += 1;
      const want = String(n);
      return o.slNo === want ? o : { ...o, slNo: want };
    });
  };
  useEffect(() => {
    if (operationsNeedRenumber(operations)) setOperations(renumberOperations(operations));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [operations]);
  const { headings: headingLibrary, addHeading: saveHeadingToLibrary } = useObHeadingLibrary();
  const { entries: subEntryLibrary, addEntry: saveSubEntry, addEntries: saveSubEntries, updateEntry: updateSubEntry, deleteEntry: deleteSubEntry, deleteEntries: deleteSubEntries } = useObSubEntryLibrary();
  const { setPhoto: saveGroupPhoto, getPhoto: getGroupPhoto } = useObGroupPhotoLibrary();
  const addBlankSubEntry = () => saveSubEntry({ heading: "", name: "", section: "sewing", machine: "", sam: "", manpower: "1", opCount: "", hpCount: "" });
  const addBlankSubEntryToHeading = (headingName) => saveSubEntry({ heading: headingName, name: "", section: "sewing", machine: "", sam: "", manpower: "1", opCount: "", hpCount: "" });
  // Insert a saved Sub Entry into this style's operations. If it's linked to a Heading, drop it at
  // the end of that heading's group (creating the heading here too, if this style doesn't have it
  // yet); otherwise just append it like a normal Add Operation.
  const insertSubEntry = (entry) => {
    const nextNo = operations.filter((o) => !o.isHeading).length + 1;
    const newOp = {
      ...blankOBOperation(),
      name: entry.name, section: entry.section, machine: entry.machine, sam: entry.sam,
      manpower: entry.manpower, opCount: entry.opCount || "", hpCount: entry.hpCount || "",
    };
    const headingName = (entry.heading || "").trim();
    if (!headingName) {
      setOperations([...operations, { ...newOp, slNo: String(nextNo) }]);
      return;
    }
    let headingIdx = -1;
    operations.forEach((o, i) => {
      if (o.isHeading && (o.name || "").trim().toLowerCase() === headingName.toLowerCase()) headingIdx = i;
    });
    if (headingIdx === -1) {
      setOperations([...operations, { ...blankOBOperation(), isHeading: true, name: headingName, slNo: "" }, { ...newOp, slNo: String(nextNo) }]);
      return;
    }
    let insertAt = operations.length;
    for (let i = headingIdx + 1; i < operations.length; i++) {
      if (operations[i].isHeading) { insertAt = i; break; }
    }
    const prevNo = num(operations[insertAt - 1]?.slNo);
    const next = [...operations];
    next.splice(insertAt, 0, { ...newOp, slNo: prevNo > 0 ? String(prevNo + 1) : String(nextNo) });
    setOperations(next);
  };
  // Insert every saved Sub Entry under one Heading into this style in a single shot — creating the
  // heading here too if this style doesn't have it yet — so a whole group (e.g. THUMB HOLE with all
  // its rows) drops in with one tap instead of inserting each row one by one.
  const insertSubEntryGroup = (headingName, rows) => {
    const cleanHeading = (headingName || "").trim();
    const cleanRows = (rows || []).filter((r) => (r.name || "").trim());
    if (!cleanRows.length) return;
    const newOps = cleanRows.map((entry) => ({
      ...blankOBOperation(),
      name: entry.name, section: entry.section, machine: entry.machine, sam: entry.sam,
      manpower: entry.manpower, opCount: entry.opCount || "", hpCount: entry.hpCount || "",
    }));
    if (!cleanHeading) {
      let nextNo = operations.filter((o) => !o.isHeading).length + 1;
      setOperations([...operations, ...newOps.map((o) => ({ ...o, slNo: String(nextNo++) }))]);
      return;
    }
    let headingIdx = -1;
    operations.forEach((o, i) => {
      if (o.isHeading && (o.name || "").trim().toLowerCase() === cleanHeading.toLowerCase()) headingIdx = i;
    });
    let base = operations;
    let insertAt;
    if (headingIdx === -1) {
      base = [...operations, { ...blankOBOperation(), isHeading: true, name: cleanHeading, slNo: "" }];
      insertAt = base.length;
    } else {
      insertAt = base.length;
      for (let i = headingIdx + 1; i < base.length; i++) {
        if (base[i].isHeading) { insertAt = i; break; }
      }
    }
    const prevNo = num(base[insertAt - 1]?.slNo);
    let nextNo = prevNo > 0 ? prevNo + 1 : base.filter((o) => !o.isHeading).length + 1;
    const withNo = newOps.map((o) => ({ ...o, slNo: String(nextNo++) }));
    const next = [...base];
    next.splice(insertAt, 0, ...withNo);
    setOperations(next);
  };
  const deleteSubEntryGroup = (headingName) => {
    const clean = (headingName || "").trim().toLowerCase();
    const ids = subEntryLibrary.filter((e) => (e.heading || "").trim().toLowerCase() === clean).map((e) => e.id);
    deleteSubEntries(ids);
  };
  // Insert a single Saved Group entry at the TOP of its Heading's section (right under the heading
  // row) — creating the heading here too if this style doesn't have it yet. Used by the "Select"
  // filter picker so every tap drops the new row above whatever is already there, pushing the
  // existing rows down, and the picker stays open for more picks.
  const insertSubEntryTop = (headingName, entry) => {
    const cleanHeading = (headingName || "").trim();
    const newOp = {
      ...blankOBOperation(),
      name: entry.name, section: entry.section, machine: entry.machine, sam: entry.sam,
      manpower: entry.manpower, opCount: entry.opCount || "", hpCount: entry.hpCount || "",
    };
    const nextNo = operations.filter((o) => !o.isHeading).length + 1;
    let headingIdx = -1;
    operations.forEach((o, i) => {
      if (o.isHeading && (o.name || "").trim().toLowerCase() === cleanHeading.toLowerCase()) headingIdx = i;
    });
    if (headingIdx === -1) {
      setOperations([...operations, { ...blankOBOperation(), isHeading: true, name: cleanHeading, slNo: "" }, { ...newOp, slNo: String(nextNo) }]);
      return;
    }
    const next = [...operations];
    next.splice(headingIdx + 1, 0, { ...newOp, slNo: String(nextNo) });
    setOperations(next);
  };
  // True if this style's Operations list already has a row with this name under this heading —
  // drives the check-mark in the Saved Group filter picker.
  const isEntryInHeading = (headingName, entryName) => {
    const h = (headingName || "").trim().toLowerCase();
    const n = (entryName || "").trim().toLowerCase();
    let currentHeading = "";
    for (const o of operations) {
      if (o.isHeading) { currentHeading = (o.name || "").trim().toLowerCase(); continue; }
      if (currentHeading === h && (o.name || "").trim().toLowerCase() === n) return true;
    }
    return false;
  };
  // True if this heading has a matching Saved Group in the Sub Entry library — drives whether the
  // quick-load tick box next to the Heading tick box shows up at all.
  const headingHasSavedGroup = (headingName) => {
    const h = (headingName || "").trim().toLowerCase();
    if (!h) return false;
    return subEntryLibrary.some((e) => (e.heading || "").trim().toLowerCase() === h);
  };
  // Tick box right next to the Heading tick box on a heading row — ticking it drops every Saved
  // Group row for this heading in right under it, pushing whatever is already there down. Same
  // top-of-section placement as the Saved Groups "Select" picker.
  const applySubEntriesToHeadingRow = (headingOp) => {
    const cleanHeading = (headingOp.name || "").trim().toLowerCase();
    if (!cleanHeading) return;
    const rows = subEntryLibrary.filter((e) => (e.heading || "").trim().toLowerCase() === cleanHeading);
    if (!rows.length) return;
    const newOps = rows.map((entry) => ({
      ...blankOBOperation(),
      name: entry.name, section: entry.section, machine: entry.machine, sam: entry.sam,
      manpower: entry.manpower, opCount: entry.opCount || "", hpCount: entry.hpCount || "",
    }));
    const idx = operations.findIndex((o) => o.id === headingOp.id);
    if (idx === -1) return;
    let nextNo = operations.filter((o) => !o.isHeading).length + 1;
    const withNo = newOps.map((o) => ({ ...o, slNo: String(nextNo++) }));
    const next = [...operations];
    next.splice(idx + 1, 0, ...withNo);
    setOperations(next);
  };
  // When OP/HP counts are used on a Sub Entry (Sewing), keep Manpower synced to their sum — same
  // behaviour as the main Operations list.
  const updateSubEntryOpHpQc = (id, patch) => {
    const entry = subEntryLibrary.find((e) => e.id === id);
    if (!entry) return;
    const merged = { ...entry, ...patch };
    const total = num(merged.opCount) + num(merged.hpCount);
    updateSubEntry(id, { ...patch, manpower: total > 0 ? String(total) : entry.manpower });
  };
  // Bulk table entry — type out a whole heading's operations (like a real Operation Breakdown
  // sheet: Sl No / Operation / Section / Machine / SAM / Op / Hp) and save them all in one shot,
  // instead of adding sub entries one at a time.
  const updateBulkRow = (id, patch) => {
    setBulkRows((rows) => rows.map((r) => {
      if (r.id !== id) return r;
      const merged = { ...r, ...patch };
      return merged;
    }));
  };
  const addBulkRow = () => setBulkRows((rows) => [...rows, blankBulkRow()]);
  const removeBulkRow = (id) => setBulkRows((rows) => (rows.length > 1 ? rows.filter((r) => r.id !== id) : rows));
  const handleSaveBulkRows = () => {
    const filled = bulkRows.filter((r) => r.name.trim());
    if (!filled.length) return;
    saveSubEntries(
      filled.map((r) => {
        const total = num(r.opCount) + num(r.hpCount);
        return {
          heading: bulkHeading.trim(), name: r.name.trim(), section: r.section, machine: r.machine,
          sam: r.sam, manpower: total > 0 ? String(total) : "1", opCount: r.opCount, hpCount: r.hpCount,
        };
      })
    );
    if (bulkHeading.trim() && bulkPhotoRef) saveGroupPhoto(bulkHeading.trim(), bulkPhotoRef);
    setSavedMsg(`Saved ${filled.length} sub entry(ies) under "${bulkHeading.trim() || "(no heading)"}" ✓`);
    setTimeout(() => setSavedMsg(""), 3000);
    setBulkHeading("");
    setBulkPhotoRef("");
    setBulkRows([blankBulkRow()]);
  };
  const handleSubExportExcel = () => {
    downloadXlsx(
      "sub_entries.xlsx",
      "Sub Entries",
      [
        ["Heading", (o) => o.heading], ["Operation", (o) => o.name], ["Section", (o) => o.section],
        ["Machine", (o) => o.machine], ["SAM", (o) => o.sam], ["Manpower", (o) => o.manpower],
        ["OP", (o) => o.opCount], ["HP", (o) => o.hpCount],
      ],
      subEntryLibrary
    );
  };
  // Sample template — one row per Sub Entry. Leave Heading blank for entries not linked to any heading.
  const handleSubDownloadTemplate = () => {
    downloadXlsx(
      "sub_entries_template.xlsx",
      "Sub Entries",
      [
        ["Heading", (o) => o.heading], ["Operation", (o) => o.name], ["Section", (o) => o.section],
        ["Machine", (o) => o.machine], ["SAM", (o) => o.sam], ["Manpower", (o) => o.manpower],
        ["OP", (o) => o.opCount], ["HP", (o) => o.hpCount],
      ],
      [
        { heading: "THUMB HOLE", name: "Loading front and back", section: "packing", machine: "ALL", sam: "0.25", manpower: "1", opCount: "", hpCount: "1" },
        { heading: "THUMB HOLE", name: "1st shoulder attach", section: "sewing", machine: "4T O/L", sam: "0.2", manpower: "1", opCount: "1", hpCount: "" },
        { heading: "THUMB HOLE", name: "Neck bindling", section: "sewing", machine: "3T O/L", sam: "0.4", manpower: "2", opCount: "2", hpCount: "" },
        { heading: "THUMB HOLE", name: "Neck top stitch", section: "sewing", machine: "2T F/L", sam: "0.25", manpower: "1", opCount: "1", hpCount: "" },
      ]
    );
  };
  const handleSubImportExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readXlsxFile(file, (rows) => {
      const imported = rows
        .map((r) => ({
          heading: String(r["Heading"] ?? "").trim(),
          name: String(r["Operation"] ?? r["Operation Name"] ?? "").trim(),
          section: String(r["Section"] ?? "sewing").toLowerCase(),
          machine: String(r["Machine"] ?? ""),
          sam: String(r["SAM"] ?? ""),
          manpower: String(r["Manpower"] ?? "1"),
          opCount: String(r["OP"] ?? ""),
          hpCount: String(r["HP"] ?? ""),
        }))
        .filter((r) => r.name);
      saveSubEntries(imported);
      setSavedMsg(`Imported ${imported.length} sub entry(ies) from Excel ✓`);
      setTimeout(() => setSavedMsg(""), 3000);
    });
    e.target.value = "";
  };
  const addOperation = () => {
    const nextNo = operations.filter((o) => !o.isHeading).length + 1;
    setOperations([...operations, { ...blankOBOperation(), slNo: String(nextNo) }]);
  };
  const addHeading = () => setOperations([...operations, { ...blankOBOperation(), isHeading: true, slNo: "" }]);
  // Insert a new blank operation right after a given row — lets the planner add a missed step
  // between two existing operations instead of only appending at the end.
  const insertOperationAfter = (id) => {
    const idx = operations.findIndex((o) => o.id === id);
    if (idx === -1) return;
    const prevNo = num(operations[idx].slNo);
    const next = [...operations];
    next.splice(idx + 1, 0, { ...blankOBOperation(), slNo: prevNo > 0 ? String(prevNo + 1) : "" });
    setOperations(next);
  };
  const updateOperation = (id, patch) => setOperations(operations.map((o) => (o.id === id ? { ...o, ...patch } : o)));
  const deleteOperation = (id) => setOperations(operations.filter((o) => o.id !== id));
  // Pressing Enter in any Sl No/Operation/Section/Machine/SAM/Manpower field jumps straight down
  // to the same field in the next row, instead of the browser's default (submit / do nothing) —
  // makes typing a long Operation Bulletin feel like filling a spreadsheet.
  const handleOpEnterKey = (e, opId, col) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    const idx = operations.findIndex((o) => o.id === opId);
    if (idx === -1) return;
    for (let i = idx + 1; i < operations.length; i++) {
      const nextOp = operations[i];
      if (nextOp.isHeading && col !== "name") continue; // heading rows only expose a Name field
      const el = document.querySelector(`[data-op-row="${nextOp.id}"][data-op-col="${col}"]`);
      if (el) { el.focus(); if (el.select) el.select(); return; }
    }
    // No next row to jump to. If Enter was pressed in the Manpower field of the very last row,
    // automatically add a new blank operation row and move focus into its Operation Name field —
    // so typing a long Operation Bulletin never has to stop to tap "Add Operation".
    if (col === "manpower") {
      const nextNo = operations.filter((o) => !o.isHeading).length + 1;
      const newOp = { ...blankOBOperation(), slNo: String(nextNo) };
      setOperations([...operations, newOp]);
      setTimeout(() => {
        const el = document.querySelector(`[data-op-row="${newOp.id}"][data-op-col="name"]`);
        if (el) { el.focus(); if (el.select) el.select(); }
      }, 50);
    }
  };

  const obFileInputRef = useRef(null);
  const subFileInputRef = useRef(null);
  const masterTemplateInputRef = useRef(null);
  const { masterTemplate, save: saveMasterTemplate } = useObMasterTemplate();
  const [masterTemplateMsg, setMasterTemplateMsg] = useState("");
  const [masterTemplateError, setMasterTemplateError] = useState("");
  // Reads the uploaded factory Excel (e.g. "MB OB SHEET.xlsx"), confirms it's a valid workbook by
  // listing its sheet names, then stores it as base64 in localStorage so it stays on this device
  // for everyone using this app to reference/reuse — until replaced or removed.
  const handleUploadMasterTemplate = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setMasterTemplateError("");
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const buffer = ev.target.result;
        const wb = XLSX.read(buffer, { type: "array" });
        saveMasterTemplate({
          fileName: file.name,
          uploadedAt: Date.now(),
          sheetNames: wb.SheetNames,
          dataBase64: arrayBufferToBase64(buffer),
        });
        setMasterTemplateMsg(`"${file.name}" upload ஆச்சு ✓`);
        setTimeout(() => setMasterTemplateMsg(""), 3000);
      } catch (err) {
        setMasterTemplateError("இந்த file-ஐ படிக்க முடியலை. சரியான Excel (.xlsx/.xls) file-ஐ upload பண்ணுங்க.");
      }
    };
    reader.onerror = () => setMasterTemplateError("File-ஐ படிக்க முடியலை. மறுபடியும் முயற்சி பண்ணுங்க.");
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };
  const handleRemoveMasterTemplate = () => {
    if (!masterTemplate) return;
    if (!confirm(`"${masterTemplate.fileName}" template-ஐ நீக்கணுமா?`)) return;
    saveMasterTemplate(null);
  };
  const handleDownloadMasterTemplate = () => {
    if (!masterTemplate) return;
    const buffer = base64ToArrayBuffer(masterTemplate.dataBase64);
    const blob = new Blob([buffer], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = masterTemplate.fileName || "OB_Template.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  };
  // Example of the layout an "OB TEMPLATE" file (like the factory's MB OB SHEET) is expected to
  // have — style header block, operation breakdown table, and a costing summary block at the
  // bottom — so Maha knows what to prepare/upload here. Purely illustrative sample values.
  const handleDownloadSampleObTemplate = () => {
    const aoa = [
      ["MB OB SHEET"],
      [],
      ["Style No", "S_116484", "", "Buyer", "QVC", "", "Description", "C Neck Long Sleeve"],
      ["Order Qty", "12000", "", "Season", "SS26", "", "Size", "S-XXL"],
      ["Target Qty / Day", "500", "", "Line Efficiency %", "65", "", "Shift Minutes", "480"],
      [],
      ["Sl No", "Operation", "Section", "Machine", "SAM", "Manpower", "OP", "HP"],
      ["", "FRONT PART", "", "", "", "", "", ""],
      [1, "Shoulder attach", "sewing", "O/L", 0.35, 1, "", ""],
      [2, "Neck binding", "sewing", "SNLS", 0.42, 1, "", ""],
      [3, "Sleeve attach", "sewing", "O/L", 0.48, 1, "", ""],
      ["", "FINAL", "", "", "", "", "", ""],
      [4, "Final check", "qc", "End Line", 0.3, 1, "", ""],
      [5, "Iron & fold", "packing", "Table", 0.25, 1, "", ""],
      [],
      ["Total SMV", 1.8, "", "Total Manpower", 5, "", "Avg Target / Hr", 108],
      ["Plan Eff %", "65%", "", "Plan OWE %", "35%", "", "Pitch Time (min)", 0.36],
      ["CM / DZN $", 1.85, "", "CM Realization", 1.72, "", "CM / M/C", 0.31],
      ["CPM (Cost / Min $)", 0.045, "", "", "", "", "", ""],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }];
    ws["!cols"] = [{ wch: 14 }, { wch: 20 }, { wch: 4 }, { wch: 16 }, { wch: 10 }, { wch: 12 }, { wch: 6 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MB OB SHEET");
    XLSX.writeFile(wb, "OB_TEMPLATE_sample.xlsx");
  };
  const { templates: obTemplates, save: saveObTemplates } = useObTemplateLibrary();
  const [selectedTemplateId, setSelectedTemplateId] = useState("");
  const handleSaveAsTemplate = () => {
    const desc = (style.obDescription || "").trim();
    if (!desc) { alert('Template save பண்ண முதலில் "Style Description" fill பண்ணுங்க.'); return; }
    const realOps = operations.filter((o) => !o.isHeading);
    if (realOps.length === 0) { alert("Save பண்ண Operation Bulletin-ல ஒரு operation கூட இல்ல."); return; }
    const cloned = operations.map((op) => ({ ...op, id: uid() }));
    const existing = obTemplates.find((t) => t.description.trim().toLowerCase() === desc.toLowerCase());
    const next = existing
      ? obTemplates.map((t) => (t.id === existing.id ? { ...t, operations: cloned, updatedAt: Date.now() } : t))
      : [...obTemplates, { id: uid(), description: desc, operations: cloned, updatedAt: Date.now() }];
    saveObTemplates(next);
    setSelectedTemplateId(existing ? existing.id : next[next.length - 1].id);
    setSavedMsg(`"${desc}" template ஆ save ஆச்சு ✓`);
    setTimeout(() => setSavedMsg(""), 3000);
  };
  const handleLoadTemplate = () => {
    const tpl = obTemplates.find((t) => t.id === selectedTemplateId);
    if (!tpl) return;
    const currentCount = operations.filter((o) => !o.isHeading).length;
    if (currentCount > 0 && !confirm(`இப்போ இருக்கிற ${currentCount} operation(s)-ஐ replace பண்ணி "${tpl.description}" template-ல இருந்து load பண்ணலாமா?`)) return;
    const cloned = tpl.operations.map((op) => ({ ...op, id: uid() }));
    onUpdateStyle({ ...style, obOperations: cloned, obDescription: (style.obDescription || "").trim() || tpl.description });
    setSavedMsg(`"${tpl.description}" template load ஆச்சு ✓`);
    setTimeout(() => setSavedMsg(""), 3000);
  };
  const handleDeleteTemplate = (id) => {
    const tpl = obTemplates.find((t) => t.id === id);
    if (!tpl || !confirm(`"${tpl.description}" template-ஐ நீக்கணுமா?`)) return;
    saveObTemplates(obTemplates.filter((t) => t.id !== id));
    if (selectedTemplateId === id) setSelectedTemplateId("");
  };
  const handleObSave = () => {
    // Data auto-saves on every edit already — this just gives an explicit confirmation.
    // Also stamps Compiled Date (once, first save) and Last Change Date (every save) for the
    // Product Costing sheet.
    const today = new Date().toLocaleDateString("en-GB");
    onUpdateStyle({ ...style, obCompiledDate: style.obCompiledDate || today, obLastChangeDate: today });
    setSavedMsg("Saved ✓");
    setTimeout(() => setSavedMsg(""), 2000);
  };
  // Builds only the sheets ticked in exportSheetSel — shared by Export Excel, Preview, and Print so
  // all three always show exactly the same content. chartImageUri is only passed at Excel-export
  // time (a pre-rasterised PNG of the Line Balancing Graph); Preview/Print omit it and get the live
  // <svg> embedded straight into the Operation Bulletin sheet's HTML instead.
  const buildSelectedObSheets = (chartImageUri = "") => {
    const rampUpExtras = {
      actualMp100: rampUpTotalMp,
      avgEfficiencyPct: rampUpAvgEffPct,
      noOfLines,
      perLineTarget: rampPerLineTarget,
      totalLineTarget: rampTotalTarget,
      totalMp: rampUpTotalMp,
      sewingSam: rampUpSam,
      shiftMinutes,
    };
    const sheets = [];
    if (exportSheetSel.ob) {
      sheets.push({ name: "Operation Bulletin", html: buildOperationBulletinReportHtml(style, operations, summary, mpBalance, balancing, machineUsage, helperSummary, includeHeadingsInExport, chartImageUri) });
    }
    if (exportSheetSel.rampup) {
      sheets.push({ name: "Ramp-up Plan", html: buildRampUpReportHtml(style, rampUpPlan, rampUpExtras, includeHeadingsInExport) });
    }
    if (exportSheetSel.costing) {
      sheets.push({ name: "Product Costing", html: buildProductCostingReportHtml(style, summary, balancing, rampUpPlan, achievableQtyAtActualMp, includeHeadingsInExport) });
    }
    return sheets;
  };
  const [obExporting, setObExporting] = useState(false);
  const handleObExportExcel = async () => {
    setObExporting(true);
    try {
      // Rasterise the Line Balancing Graph to PNG first — ExcelJS can only embed jpeg/png/gif
      // pictures, not the live <svg> that Preview/Print use directly.
      const graph = buildLineBalancingGraphSvg(balancing);
      const chartImageUri = graph.svg ? await svgToPngDataUri(graph.svg, graph.width, graph.height) : "";
      const sheets = buildSelectedObSheets(chartImageUri);
      if (sheets.length === 0) return;
      await downloadHtmlAsExcel(`${style.styleNo || "style"}_operation_bulletin.xlsx`, sheets);
    } finally {
      setObExporting(false);
    }
  };
  const handleObPrint = () => {
    const sheets = buildSelectedObSheets();
    if (sheets.length === 0) return;
    const win = window.open("", "_blank");
    if (!win) return;
    const body = sheets.map((sh) => `<h2 style="font-family:Calibri,Arial,sans-serif;">${escapeHtml(sh.name)}</h2><table style="border-collapse:collapse;margin-bottom:24px;">${sh.html.replace(/^<table[^>]*>|<\/table>$/g, "")}</table>`).join("");
    win.document.write(`<!DOCTYPE html><html><head><title>${escapeHtml(style.styleNo || "Operation Bulletin")}</title></head><body>${body}</body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => win.print(), 300);
  };
  // Sample template showing how Heading rows + Sl No sequence should be laid out in Excel —
  // Import reads rows top-to-bottom exactly as they appear in the sheet, so whatever order the
  // rows are in here is the order they'll be added to the Operation Bulletin. Mark a row as a
  // heading by putting Y in the Heading column (Section/Machine/SAM/Manpower are ignored for those rows).
  const handleObDownloadTemplate = () => {
    downloadXlsx(
      "operation_bulletin_template.xlsx",
      "Operation Bulletin",
      [
        ["Sl No", (o) => o.slNo], ["Heading", (o) => o.heading], ["Operation", (o) => o.name], ["Section", (o) => o.section],
        ["Machine", (o) => o.machine], ["SAM", (o) => o.sam], ["Manpower", (o) => o.manpower],
        ["OP", (o) => o.opCount], ["HP", (o) => o.hpCount],
      ],
      [
        { slNo: "", heading: "Y", name: "FRONT PART", section: "", machine: "", sam: "", manpower: "", opCount: "", hpCount: "" },
        { slNo: "1", heading: "", name: "Shoulder attach", section: "sewing", machine: "O/L", sam: "0.35", manpower: "1", opCount: "", hpCount: "" },
        { slNo: "2", heading: "", name: "Neck binding", section: "sewing", machine: "SNLS", sam: "0.42", manpower: "1", opCount: "", hpCount: "" },
        { slNo: "", heading: "Y", name: "FINAL", section: "", machine: "", sam: "", manpower: "", opCount: "", hpCount: "" },
        { slNo: "3", heading: "", name: "Final check", section: "qc", machine: "End Line", sam: "0.30", manpower: "1", opCount: "", hpCount: "" },
      ]
    );
  };
  const handleObImportExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readXlsxFile(file, (rows) => {
      let nextNo = operations.filter((o) => !o.isHeading).length;
      const imported = rows.map((r) => {
        const isHeading = /^(y|yes|true|1)$/i.test(String(r["Heading"] ?? "").trim());
        if (isHeading) {
          return { ...blankOBOperation(), isHeading: true, name: String(r["Operation"] ?? r["Operation Name"] ?? "") };
        }
        nextNo += 1;
        return {
          ...blankOBOperation(),
          slNo: String(r["Sl No"] || nextNo),
          name: String(r["Operation"] ?? r["Operation Name"] ?? ""),
          section: String(r["Section"] ?? "sewing").toLowerCase(),
          machine: String(r["Machine"] ?? ""),
          sam: String(r["SAM"] ?? ""),
          manpower: String(r["Manpower"] ?? "1"),
          opCount: String(r["OP"] ?? ""), hpCount: String(r["HP"] ?? ""),
        };
      });
      setOperations([...operations, ...imported]);
      setSavedMsg(`Imported ${imported.length} row(s) from Excel ✓`);
      setTimeout(() => setSavedMsg(""), 3000);
    });
    e.target.value = "";
  };

  // When OP/HP counts are used, keep Manpower (drives all the capacity/line-balancing calcs) synced to their sum.
  const updateOpHpQc = (id, patch) => {
    const op = operations.find((o) => o.id === id);
    if (!op) return;
    const merged = { ...op, ...patch };
    const total = num(merged.opCount) + num(merged.hpCount);
    updateOperation(id, { ...patch, manpower: total > 0 ? String(total) : op.manpower });
  };
  const machineUsage = useMemo(() => computeMachineUsage(operations), [operations]);
  // Saved Sub Entries grouped by Heading — drives the "Insert as Group" chips so a whole set
  // (e.g. THUMB HOLE with all its rows) can be inserted or deleted together in one tap.
  const subEntryGroups = useMemo(() => {
    const order = [];
    const map = {};
    subEntryLibrary.forEach((entry) => {
      const heading = (entry.heading || "").trim();
      if (!heading) return;
      const key = heading.toLowerCase();
      if (!map[key]) { map[key] = { heading, rows: [] }; order.push(key); }
      map[key].rows.push(entry);
    });
    return order.map((key) => map[key]);
  }, [subEntryLibrary]);
  const [openGroupKey, setOpenGroupKey] = useState("");
  const [groupFilterText, setGroupFilterText] = useState("");
  const [editGroupKey, setEditGroupKey] = useState("");
  const helperSummary = useMemo(() => computeHelperSummary(operations), [operations]);
  // Whichever departments actually have operations entered (e.g. only Sewing + QC filled) —
  // drives the auto combined-SMV line under Section-wise Usage Summary (e.g. "Sewing + QC = 7.1").
  const filledSections = useMemo(() => OB_SECTIONS.filter((s) => summary.sections[s.key].count > 0), [summary]);
  const filledSectionsSmv = useMemo(() => filledSections.reduce((a, s) => a + summary.sections[s.key].smv, 0), [filledSections, summary]);

  const [headEffPct, setHeadEffPct] = useState("");
  const [opsView, setOpsView] = useState("table"); // 'cards' | 'table' — Operations list layout
  // Drag-to-resize column widths for the Operation Bulletin table (Table view). Widths live only in
  // this component's state (reset on reload) — the goal is letting the user widen the Operation /
  // Heading name column (or any other) to see long text, not persisting a saved layout.
  const OB_COL_DEFAULTS = { slno: 44, head: 60, operation: 200, section: 96, machine: 128, sam: 64, manpower: 116, req: 96 };
  const [obColWidths, setObColWidths] = useState(OB_COL_DEFAULTS);
  const obColDrag = useRef(null);
  const startObColResize = (key) => (e) => {
    e.preventDefault();
    obColDrag.current = { key, startX: e.clientX, startWidth: obColWidths[key] };
    const onMove = (ev) => {
      if (!obColDrag.current) return;
      const { key: k, startX, startWidth } = obColDrag.current;
      const next = Math.max(40, startWidth + (ev.clientX - startX));
      setObColWidths((w) => ({ ...w, [k]: next }));
    };
    const onUp = () => {
      obColDrag.current = null;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };
  const ObColResizeHandle = ({ colKey }) => (
    <span
      onMouseDown={startObColResize(colKey)}
      title="Drag to resize column"
      className="absolute right-0 top-0 h-full w-2 cursor-col-resize select-none hover:bg-amber-300/60 active:bg-amber-400"
    />
  );
  // Working Hours/Day drives every shift-based calc below — defaults to the standard 8-hour
  // shift (SHIFT_MINUTES) when the field is left blank.
  const shiftMinutes = num(style?.obWorkingHours) > 0 ? num(style.obWorkingHours) * 60 : SHIFT_MINUTES;
  const mpBalance = useMemo(
    () => computeManpowerBalance(operations, style?.obTargetQtyPerDay, headEffPct || style?.targetEfficiencyPct, shiftMinutes),
    [operations, style?.obTargetQtyPerDay, headEffPct, style?.targetEfficiencyPct, shiftMinutes]
  );
  const sewingSam = summary.sections.sewing.smv;
  const suggestedMp = useMemo(
    () => suggestedManpower(style?.obTargetQtyPerDay, sewingSam, style?.targetEfficiencyPct, shiftMinutes),
    [style?.obTargetQtyPerDay, sewingSam, style?.targetEfficiencyPct, shiftMinutes]
  );
  // Manpower needed to hit the customer target at a straight 100% efficiency (ignores the
  // Target Efficiency % field) — the "full target" manpower figure.
  const suggestedMp100 = useMemo(
    () => suggestedManpower(style?.obTargetQtyPerDay, sewingSam, 100, shiftMinutes),
    [style?.obTargetQtyPerDay, sewingSam, shiftMinutes]
  );
  // Total manpower allocated across Sewing operations in this OB (sum of each op's Manpower
  // field) — the fallback source for Actual Sewing Manpower when that field is left blank.
  const totalAllocatedSewingMp = useMemo(
    () => operations.filter((op) => op.section === "sewing" && !op.isHeading).reduce((a, op) => a + allocatedMpOf(op), 0),
    [operations]
  );
  const actualSewingMpManual = style?.obActualSewingMp !== undefined && style?.obActualSewingMp !== "";
  const actualSewingMp = actualSewingMpManual ? num(style.obActualSewingMp) : totalAllocatedSewingMp;
  const mpFulfilledPct = suggestedMp > 0 ? (actualSewingMp / suggestedMp) * 100 : 0;
  // Total manpower across the WHOLE OB — Cutting + Sewing + QC + Packing combined (every non-heading
  // operation row), not Sewing alone. This feeds the Ramp-up Plan section and CM/Machine, since a
  // day's real achievable output depends on every department's headcount, not just sewing operators.
  const totalObManpower = operations.filter((op) => !op.isHeading).reduce((a, op) => a + allocatedMpOf(op), 0);
  // Reverse calc: with the actual Sewing manpower (entered above, or the Total Allocated Manpower
  // fallback), what qty/day is achievable at 100% efficiency — base line-planning off the real
  // headcount instead of the customer's target.
  const achievableQtyAtActualMp = useMemo(() => {
    if (actualSewingMp <= 0 || sewingSam <= 0) return 0;
    return (actualSewingMp * shiftMinutes) / sewingSam;
  }, [actualSewingMp, sewingSam, shiftMinutes]);
  // No. of Lines the order is planned across, and the resulting per-line day target — feeds both
  // the Ramp-up Plan section and the Product Costing sheet's "Planned Line" field.
 const noOfLines = num(style?.obPlannedLines || 0);
  const linePerDayTarget = noOfLines > 0 ? num(style.obTargetQtyPerDay) / noOfLines : 0;
  // Ramp-up Plan uses its own Per-Line Target Qty/Day (obRampPerLineTarget) so the planner can key
  // in the line's real takt (e.g. 1200 pcs/line/day) directly instead of back-deriving it from the
  // Customer Target Qty/Day above. Falls back to the derived linePerDayTarget when left blank, so
  // existing styles keep behaving exactly as before. Total daily capacity across every line
  // (perLineTarget × noOfLines) is what actually drives the ramp-up day-count / lead-time maths —
  // this is what fixes the lead-time so it matches the per-line plan (e.g. 1200 × 2 lines =
  // 2400/day; 44400 remaining ÷ 2400 = 18.5 days + 12 ramp-up days = 30.5 days lead time).
  const rampPerLineTarget = num(style?.obRampPerLineTarget || 0) > 0 ? num(style?.obRampPerLineTarget || 0) : linePerDayTarget;
const rampTotalTarget = noOfLines > 0 ? rampPerLineTarget * noOfLines : num(style?.obTargetQtyPerDay || 0);
  // The Operation Bulletin's manpower entries (totalObManpower) describe ONE line's headcount.
  // With 2+ lines running the same operations side by side, the manpower actually on the floor is
  // that many times bigger — so Ramp-up's efficiency maths must multiply totalObManpower by
  // No. of Lines too, or Eff% comes out wrong (target qty scales with lines, manpower didn't).
  const rampUpTotalMp = totalObManpower * (noOfLines > 0 ? noOfLines : 1);
  const rampUpPlan = useMemo(
    () => computeRampUpPlan(style?.orderQty, rampTotalTarget, style?.obRampUpDays, noOfLines),
    [style?.orderQty, rampTotalTarget, style?.obRampUpDays, noOfLines]
  );
  // Combined SAM for Ramp-up efficiency calc — Sewing + QC + Packing (summary.sewToPackSmv, same
  // figure as the "Sewing + QC + Packing" total on the Section-wise Usage Summary card), not
  // Sewing SAM alone, since Ramp-up planning is against the full garment work content.
  const rampUpSam = summary.sewToPackSmv;
  const setRampUpDays = (rows) => onUpdateStyle({ ...style, obRampUpDays: rows });
  const addRampUpDay = () => setRampUpDays([...(style.obRampUpDays || []), { id: uid(), qty: "" }]);
  const updateRampUpDay = (id, qty) => setRampUpDays((style.obRampUpDays || []).map((r) => (r.id === id ? { ...r, qty } : r)));
  const removeRampUpDay = (id) => setRampUpDays((style.obRampUpDays || []).filter((r) => r.id !== id));
  // Average Eff % across sewing operations (Plan Cap ÷ Manpower Allocated) — feeds the CM calculation.
  const avgSewingEffPct = useMemo(() => {
    // mpBalance now spans every section (Cutting/Sewing/QC/Packing) so its columns can show up
    // on those rows too — filter back to Sewing-only here since this figure specifically feeds the
    // CM calculation, a sewing-line figure.
    const rows = (mpBalance.rows || []).filter((r) => r.section === "sewing");
    if (rows.length === 0) return 0;
    return rows.reduce((a, r) => a + (r.effPct || 0), 0) / rows.length;
  }, [mpBalance]);
  // Per-day production efficiency — day capacity @ 100% = (Total Manpower × Working Minutes) ÷ SAM;
  // that day's Eff% = day qty ÷ that capacity × 100 (same as qty × SAM ÷ (Manpower × Minutes) × 100).
  const rampDayEffPct = useMemo(() => {
    if (rampUpTotalMp <= 0 || rampUpSam <= 0) return () => 0;
    const capacity100 = (rampUpTotalMp * shiftMinutes) / rampUpSam;
    return (qty) => (capacity100 > 0 ? (num(qty) / capacity100) * 100 : 0);
  }, [rampUpTotalMp, rampUpSam, shiftMinutes]);
  // Ramp-up "Avg Efficiency %" — average of every day's production efficiency (above), across the
  // WHOLE ramp-up plan (manual entries + auto-filled days) — the total-days average, not a single
  // ratio-of-totals figure.
  const rampUpAvgEffPct = useMemo(() => {
    if (rampUpPlan.rows.length === 0) return 0;
    const sum = rampUpPlan.rows.reduce((a, r) => a + rampDayEffPct(r.qty), 0);
    return sum / rampUpPlan.rows.length;
  }, [rampUpPlan.rows, rampDayEffPct]);
  // CM (Cost of Making) calculation — CM/Minute (CPM) is a manual factory input. Formula:
  // CM = (Sew-to-Pack SMV ÷ Avg Efficiency %) × CPM — per piece; × 12 for per dozen. Avg Efficiency
  // uses a dedicated manual entry (obCmAvgEffPct) when set, else falls back to the Avg Efficiency %
  // already computed on the Ramp-up sheet from the sewing operations. Currency and Per Pcs / Per Dzn
  // basis are both selectable per style; CM Realization % is now a direct manual entry (factory's
  // own realization figure) instead of being force-derived from Plan Eff%.
  const safeStyle = style || {};

  const cmCurrency = safeStyle.obCmCurrency || "USD";
  const cmSymbol = cmCurrencySymbol(cmCurrency);
  const cmBasis = safeStyle.obCmBasis === "pcs" ? "pcs" : "dzn";
  const cmPerMinute = num(safeStyle.obCmPerMinute || 0);
  const cmAvgEffPct = num(safeStyle.obCmAvgEffPct || 0) > 0 ? num(safeStyle.obCmAvgEffPct || 0) : (avgSewingEffPct || 100);
  const cmEffFactor = cmAvgEffPct > 0 ? cmAvgEffPct / 100 : 1;
  const cmValuePerPc = cmPerMinute > 0 && cmEffFactor > 0 ? (summary.sewToPackSmv / cmEffFactor) * cmPerMinute : 0;
  const cmValuePerDzn = cmValuePerPc * 12;
  const cmPerDzn = cmBasis === "pcs" ? cmValuePerPc : cmValuePerDzn;
  const cmPerMachine = totalObManpower > 0 ? cmPerDzn / totalObManpower : 0;  const [includeHeadingsInExport, setIncludeHeadingsInExport] = useState(true);
  // Which report sheets go into the Excel export — planner can untick Ramp-up or Product Costing
  // if this style doesn't need them, instead of always getting all 3 sheets bundled together.
  const [exportSheetSel, setExportSheetSel] = useState({ ob: true, rampup: true, costing: true });
  const toggleExportSheet = (key) => setExportSheetSel((prev) => ({ ...prev, [key]: !prev[key] }));
  const [showExportPreview, setShowExportPreview] = useState(false);
  useEffect(() => {
    if (!showExportPreview) return;
    const prevOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = prevOverflow; };
  }, [showExportPreview]);

  const syncToStyleSMV = () => {
    const nextDepartments = (style.departments || blankDepartments()).map((d) => {
      if (d.key === "cutting") return { ...d, smv: summary.sections.cutting.smv ? String(summary.sections.cutting.smv) : d.smv };
      if (d.key === "sewing") return { ...d, smv: summary.sewToQcSmv ? String(summary.sewToQcSmv) : d.smv };
      if (d.key === "packing") return { ...d, smv: summary.sections.packing.smv ? String(summary.sections.packing.smv) : d.smv };
      return d;
    });
    onUpdateStyle({ ...style, departments: nextDepartments });
    setSavedMsg("Synced to Style SMV (Cutting / Sewing incl. QC / Packing)");
    setTimeout(() => setSavedMsg(""), 3000);
  };

  const chartData = balancing.rows.map((r) => ({ name: r.name || `Op ${r.slNo}`, capacity: Math.round(r.capacityPerHr * 10) / 10 }));

  return (
    <div className="space-y-4 pb-16">
      <datalist id="ob-machine-list">
        {OB_MACHINE_LIST.map((m) => <option key={m} value={m} />)}
      </datalist>
      <datalist id="ob-packing-list">
        {OB_PACKING_LIST.map((m) => <option key={m} value={m} />)}
      </datalist>
      <datalist id="ob-cutting-list">
        {OB_CUTTING_LIST.map((m) => <option key={m} value={m} />)}
      </datalist>
      <datalist id="ob-heading-list">
        {headingLibrary.map((h) => <option key={h} value={h} />)}
      </datalist>

      {obTab === "newstyle" ? (
        <div className="rounded-2xl border-2 border-dashed border-amber-300 bg-white p-4 space-y-2.5">
          <p className="text-[11px] font-bold text-amber-700 uppercase tracking-wide">New Style</p>
          <Field label="Style No" value={newStyleNo} onChange={setNewStyleNo} type="text" placeholder="e.g. S_116484" />
          <div className="grid grid-cols-2 gap-2.5">
            <Field label="Buyer" value={newBuyer} onChange={setNewBuyer} type="text" placeholder="e.g. QVC" />
            <Field label="Order Qty" value={newOrderQty} onChange={setNewOrderQty} suffix="pcs" />
          </div>
          <PhotoRefUpload value={newPhotoRef} onChange={setNewPhotoRef} label="Style Photo Reference" />
          <div className="flex gap-2">
            <button onClick={handleCreateStyle} disabled={!newStyleNo} className="flex-1 rounded-xl bg-amber-500 disabled:opacity-40 text-stone-900 font-bold py-2.5 text-sm active:bg-amber-400">
              Create Style
            </button>
            {styles.length > 0 && (
              <button onClick={() => setObTab("target")} className="rounded-xl border border-stone-200 text-stone-500 px-4 py-2.5 text-sm font-bold">
                Cancel
              </button>
            )}
          </div>
        </div>
      ) : obTab === "templateupload" ? (
        <div className="rounded-2xl border-2 border-dashed border-amber-300 bg-white p-4 space-y-3">
          <p className="text-[11px] font-bold text-amber-700 uppercase tracking-wide">OB Template Upload</p>
          <p className="text-[11px] text-stone-500">
            Factory-ல வேற எங்காவது இருக்கிற <b>OB TEMPLATE Excel file</b>-ஐ (உதா: "MB OB SHEET.xlsx") இங்க upload பண்ணுங்க — இந்த device-ல save ஆகி, எப்போ வேணா திரும்ப பாக்கலாம் / download பண்ணலாம். இது இந்த app-ல உள்ள operations data-வை பாதிக்காது.
          </p>

          <div className="rounded-xl border border-dashed border-stone-300 bg-stone-50 p-3 flex flex-wrap items-center gap-2.5">
            <span className="text-[11px] text-stone-500 flex-1 min-w-[160px]">
              File எப்படி இருக்கணும்-ன்னு தெரியலையா? Sample file-ஐ பாருங்க — Style header, Operation table, கடைசில CM/DZN$, Plan Eff%, Pitch Time மாதிரி costing summary.
            </span>
            <button onClick={handleDownloadSampleObTemplate} className="flex items-center gap-1.5 rounded-lg border border-stone-300 bg-white text-stone-600 px-3 py-2 text-xs font-bold shrink-0">
              <Download size={13} /> Sample Template
            </button>
          </div>

          <input
            ref={masterTemplateInputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={handleUploadMasterTemplate}
            className="hidden"
          />

          {masterTemplateError && (
            <p className="text-[11px] font-semibold text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2">{masterTemplateError}</p>
          )}

          {masterTemplate ? (
            <div className="rounded-xl border border-emerald-200 bg-emerald-50/60 p-3 space-y-2">
              <div className="flex items-start gap-2.5">
                <span className="h-9 w-9 rounded-lg bg-emerald-600 text-white flex items-center justify-center shrink-0">
                  <FileSpreadsheet size={17} />
                </span>
                <div className="min-w-0 flex-1">
                  <p className="text-xs font-bold text-stone-800 truncate">{masterTemplate.fileName}</p>
                  <p className="text-[10px] text-stone-500">
                    Uploaded {new Date(masterTemplate.uploadedAt).toLocaleDateString()} · {new Date(masterTemplate.uploadedAt).toLocaleTimeString()}
                  </p>
                  {masterTemplate.sheetNames?.length > 0 && (
                    <p className="text-[10px] text-stone-500 mt-0.5">Sheets: {masterTemplate.sheetNames.join(", ")}</p>
                  )}
                </div>
              </div>
              <div className="flex flex-wrap gap-2">
                <button onClick={handleDownloadMasterTemplate} className="flex items-center gap-1.5 rounded-lg border border-stone-200 bg-white text-stone-600 px-3 py-2 text-xs font-bold">
                  <Download size={13} /> Download
                </button>
                <button onClick={() => masterTemplateInputRef.current?.click()} className="flex items-center gap-1.5 rounded-lg border border-stone-200 bg-white text-stone-600 px-3 py-2 text-xs font-bold">
                  <Upload size={13} /> Replace
                </button>
                <button onClick={handleRemoveMasterTemplate} className="flex items-center gap-1.5 rounded-lg border border-red-200 bg-white text-red-600 px-3 py-2 text-xs font-bold">
                  <Trash2 size={13} /> Remove
                </button>
              </div>
            </div>
          ) : (
            <button
              onClick={() => masterTemplateInputRef.current?.click()}
              className="w-full flex flex-col items-center justify-center gap-1.5 rounded-xl border-2 border-dashed border-amber-300 bg-amber-50/50 px-4 py-6 text-amber-700"
            >
              <Upload size={22} />
              <span className="text-xs font-bold">OB Template Excel file upload பண்ண தட்டுங்க</span>
              <span className="text-[10px] text-amber-600/80">.xlsx / .xls</span>
            </button>
          )}

          {masterTemplateMsg && <span className="flex items-center gap-1 text-xs font-semibold text-emerald-600"><Check size={13} /> {masterTemplateMsg}</span>}
        </div>
      ) : obTab === "savedob" ? (
        <SavedObScreen styles={styles} onEditStyle={(id) => { setStyleId(id); setObTab("target"); }} />
      ) : obTab === "subentries" ? null : (
        <div className="rounded-2xl border border-stone-200 bg-white p-4 space-y-3">
          {styles.length > 0 ? (
            <label className="flex flex-col gap-1">
              <span className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">Switch Style</span>
              <select value={styleId} onChange={(e) => setStyleId(e.target.value)} className="rounded-xl border border-stone-200 bg-white px-3 py-2.5 text-sm outline-none focus:border-amber-500">
                {styles.map((s) => <option key={s.id} value={s.id}>{s.styleNo || "Untitled"}{s.buyer ? ` · ${s.buyer}` : ""}</option>)}
              </select>
            </label>
          ) : (
            <p className="text-xs text-stone-400">No styles yet — use the <b>New Style</b> item in the Operation Bulletin menu to create one.</p>
          )}

          {style && (
            <>
              <div className="grid grid-cols-2 gap-2.5">
                <Field label="Style No" value={style.styleNo || ""} onChange={(v) => onUpdateStyle({ ...style, styleNo: v })} type="text" placeholder="e.g. S_116484" />
                <Field label="Buyer" value={style.buyer || ""} onChange={(v) => onUpdateStyle({ ...style, buyer: v })} type="text" placeholder="e.g. QVC" />
              </div>
              <div className="grid grid-cols-3 gap-2.5">
                <Field label="Order Qty" value={style.orderQty || ""} onChange={(v) => onUpdateStyle({ ...style, orderQty: v })} suffix="pcs" />
                <Field label="Season" value={style.obSeason || ""} onChange={(v) => onUpdateStyle({ ...style, obSeason: v })} type="text" placeholder="e.g. SS26" />
                <Field label="Size" value={style.obSize || ""} onChange={(v) => onUpdateStyle({ ...style, obSize: v })} type="text" placeholder="e.g. S-XXL" />
              </div>
              <Field label="Style Description" value={style.obDescription || ""} onChange={(v) => onUpdateStyle({ ...style, obDescription: v })} type="text" placeholder="e.g. C Neck Long Sleeve" />
              <PhotoRefUpload value={style.photoRef || ""} onChange={(v) => onUpdateStyle({ ...style, photoRef: v })} label="Garment Sketch / Photo (Product Costing sheet)" />
              <PhotoGalleryUpload value={style.extraPhotoRefs || []} onChange={(v) => onUpdateStyle({ ...style, extraPhotoRefs: v })} label="Extra Photos (Product Costing & Operation Bulletin sheet-க்கும் போகும்)" />

              <div className="rounded-xl border border-stone-200 bg-stone-50/60 p-3 space-y-2">
                <p className="text-[10px] font-bold text-stone-500 uppercase tracking-wide">Product Costing Details (Excel export-ல Product Costing sheet-க்கு)</p>
                <div className="grid grid-cols-2 gap-2.5">
                  <Field label="Factory" value={style.obFactory || ""} onChange={(v) => onUpdateStyle({ ...style, obFactory: v })} type="text" placeholder="e.g. EMBEE GROUP" />
                  <Field label="Country" value={style.obCountry || ""} onChange={(v) => onUpdateStyle({ ...style, obCountry: v })} type="text" placeholder="e.g. EGYPT" />
                </div>
                <div className="grid grid-cols-3 gap-2.5">
                  <Field label="Lead Time (days)" value={style.obLeadTime || ""} onChange={(v) => onUpdateStyle({ ...style, obLeadTime: v })} />
                  <Field label="Planned Line" value={style.obPlannedLines || ""} onChange={(v) => onUpdateStyle({ ...style, obPlannedLines: v })} />
                  <Field label="Compiled By" value={style.obCompiledBy || ""} onChange={(v) => onUpdateStyle({ ...style, obCompiledBy: v })} type="text" placeholder="e.g. DINESH" />
                </div>
                {(style.obCompiledDate || style.obLastChangeDate) && (
                  <p className="text-[10px] text-stone-400">
                    {style.obCompiledDate ? `Compiled: ${style.obCompiledDate}` : ""}{style.obCompiledDate && style.obLastChangeDate ? " · " : ""}{style.obLastChangeDate ? `Last Change: ${style.obLastChangeDate}` : ""}
                  </p>
                )}
              </div>

              {obTab === "ops" && (
              <div className="rounded-xl border border-dashed border-amber-300 bg-amber-50/50 p-3 space-y-2">
                <p className="text-[10px] font-bold text-amber-700 uppercase tracking-wide">OB Template Library</p>
                <div className="flex flex-wrap items-center gap-2">
                  <select value={selectedTemplateId} onChange={(e) => setSelectedTemplateId(e.target.value)} className="flex-1 min-w-[160px] rounded-lg border border-stone-200 bg-white px-2.5 py-2 text-xs outline-none focus:border-amber-500">
                    <option value="">Select saved Style Description...</option>
                    {obTemplates.map((t) => <option key={t.id} value={t.id}>{t.description} ({t.operations.filter((o) => !o.isHeading).length} ops)</option>)}
                  </select>
                  <button onClick={handleLoadTemplate} disabled={!selectedTemplateId} className="flex items-center gap-1.5 rounded-lg bg-amber-500 text-stone-900 px-3 py-2 text-xs font-bold disabled:opacity-40 active:bg-amber-400">
                    <Download size={13} /> Load
                  </button>
                  {selectedTemplateId && (
                    <button onClick={() => handleDeleteTemplate(selectedTemplateId)} className="flex items-center gap-1.5 rounded-lg border border-red-200 text-red-600 px-2.5 py-2 text-xs font-bold">
                      <Trash2 size={13} />
                    </button>
                  )}
                </div>
                <button onClick={handleSaveAsTemplate} className="w-full flex items-center justify-center gap-1.5 rounded-lg border border-amber-300 text-amber-700 px-3 py-2 text-xs font-bold active:bg-amber-100">
                  <Save size={13} /> Save Current Operations as Template
                </button>
                <p className="text-[10px] text-stone-400">ஒரு Style Description-க்கு (ex: "C Neck Long Sleeve") operations fill பண்ணி/Excel Import பண்ணி — "Save Current Operations as Template" தட்டுங்க. அடுத்த முறை அதே மாதிரி style வந்தா, இங்க Select பண்ணி Load தட்டுங்க — Operation Bulletin முழுசா auto-ஆ வந்துடும், தேவையான மாற்றம் மட்டும் edit பண்ணி வேலையை சீக்கிரம் முடிக்கலாம்.</p>
              </div>
              )}
            </>
          )}
        </div>
      )}

      {style && obTab !== "newstyle" && obTab !== "templateupload" && obTab !== "savedob" && obTab !== "subentries" && (
        <div className="rounded-2xl border border-stone-200 bg-white p-3 flex flex-wrap items-center gap-2">
          <button onClick={handleObSave} className="flex items-center gap-1.5 rounded-xl bg-emerald-600 text-white px-3.5 py-2 text-xs font-bold active:bg-emerald-700">
            <Save size={14} /> Save
          </button>
          <button onClick={handleObExportExcel} disabled={obExporting} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold disabled:opacity-50">
            <FileSpreadsheet size={14} /> {obExporting ? "Preparing..." : "Export Excel"}
          </button>
          <button onClick={() => setShowExportPreview(true)} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold">
            <Eye size={14} /> Preview
          </button>
          <button onClick={handleObPrint} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold">
            <Printer size={14} /> Print
          </button>
          <div className="flex items-center gap-2.5 flex-wrap">
            <label className="flex items-center gap-1.5 text-[11px] font-semibold text-stone-500 select-none">
              <input type="checkbox" checked={exportSheetSel.ob} onChange={() => toggleExportSheet("ob")} className="w-3.5 h-3.5 accent-amber-500" />
              Operation Bulletin
            </label>
            <label className="flex items-center gap-1.5 text-[11px] font-semibold text-stone-500 select-none">
              <input type="checkbox" checked={exportSheetSel.rampup} onChange={() => toggleExportSheet("rampup")} className="w-3.5 h-3.5 accent-amber-500" />
              Ramp-up Plan
            </label>
            <label className="flex items-center gap-1.5 text-[11px] font-semibold text-stone-500 select-none">
              <input type="checkbox" checked={exportSheetSel.costing} onChange={() => toggleExportSheet("costing")} className="w-3.5 h-3.5 accent-amber-500" />
              Product Costing
            </label>
          </div>
          <label className="flex items-center gap-1.5 text-[11px] font-semibold text-stone-500 select-none">
            <input type="checkbox" checked={includeHeadingsInExport} onChange={(e) => setIncludeHeadingsInExport(e.target.checked)} className="w-3.5 h-3.5 accent-amber-500" />
            Heading வேணுமா (column headings-ஓட Excel export)
          </label>
          <button onClick={() => obFileInputRef.current?.click()} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold">
            <Upload size={14} /> Import Excel
          </button>
          <input ref={obFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleObImportExcel} className="hidden" />
          <button onClick={handleObDownloadTemplate} className="flex items-center gap-1.5 rounded-xl border border-dashed border-amber-300 text-amber-700 px-3.5 py-2 text-xs font-bold">
            <Download size={14} /> Download Template
          </button>
          {savedMsg && <span className="flex items-center gap-1 text-xs font-semibold text-emerald-600 ml-auto"><Check size={13} /> {savedMsg}</span>}
        </div>
      )}
      {showExportPreview && (
        <div className="fixed inset-0 z-50 bg-black/50 flex items-end sm:items-center justify-center p-0 sm:p-6">
          <div className="bg-white w-full sm:max-w-3xl sm:rounded-2xl h-[92vh] [@supports(height:100dvh)]:h-[92dvh] sm:h-[85vh] flex flex-col overflow-hidden overscroll-contain">
            <div className="flex items-center justify-between px-4 py-3 border-b border-stone-200 shrink-0">
              <h3 className="text-sm font-bold text-stone-700">Export Preview {includeHeadingsInExport ? "" : "(no headings)"}</h3>
              <button onClick={() => setShowExportPreview(false)} className="text-stone-400 active:text-stone-700"><X size={18} /></button>
            </div>
            <div className="flex-1 overflow-y-auto overscroll-contain p-4 space-y-6 bg-stone-50" style={{ WebkitOverflowScrolling: "touch" }}>
              {buildSelectedObSheets().length === 0 ? (
                <p className="text-xs text-stone-400 text-center py-8">Sheet-ஐ தேர்ந்தெடுங்க (Operation Bulletin / Ramp-up Plan / Product Costing) — preview பண்ண குறைந்தது ஒன்று வேணும்.</p>
              ) : (
                buildSelectedObSheets().map((sh) => (
                  <div key={sh.name} className="bg-white rounded-xl border border-stone-200 overflow-x-auto">
                    <h4 className="text-xs font-bold text-stone-500 uppercase tracking-wide px-3 pt-3">{sh.name}</h4>
                    {sh.name === "Operation Bulletin" && (
                      <table className="w-full text-[11px] border-collapse mt-2">
                        <thead>
                          <tr className="bg-stone-100 text-stone-500 text-left">
                            <th className="px-2 py-1.5 border border-stone-200">Sl No</th>
                            <th className="px-2 py-1.5 border border-stone-200">Operation</th>
                            <th className="px-2 py-1.5 border border-stone-200">Section</th>
                            <th className="px-2 py-1.5 border border-stone-200">Machine</th>
                            <th className="px-2 py-1.5 border border-stone-200">SAM</th>
                            <th className="px-2 py-1.5 border border-stone-200">Manpower</th>
                          </tr>
                        </thead>
                        <tbody>
                          {operations.map((op) =>
                            op.isHeading ? (
                              <tr key={op.id}>
                                <td colSpan={6} className="px-2 py-1.5 border border-stone-200 bg-amber-50 font-bold text-amber-700">{op.name}</td>
                              </tr>
                            ) : (
                              <tr key={op.id}>
                                <td className="px-2 py-1 border border-stone-200 tabular-nums">{op.slNo}</td>
                                <td className="px-2 py-1 border border-stone-200">{op.name}</td>
                                <td className="px-2 py-1 border border-stone-200 capitalize">{(OB_SECTIONS.find((x) => x.key === op.section) || {}).label || op.section}</td>
                                <td className="px-2 py-1 border border-stone-200">{op.machine}</td>
                                <td className="px-2 py-1 border border-stone-200 tabular-nums">{money(op.sam)}</td>
                                <td className="px-2 py-1 border border-stone-200 tabular-nums">{op.manpower}</td>
                              </tr>
                            )
                          )}
                        </tbody>
                      </table>
                    )}
                    <div className="p-3" dangerouslySetInnerHTML={{ __html: sh.html }} />
                  </div>
                ))
              )}
            </div>
            <div className="flex items-center gap-2 px-4 py-3 border-t border-stone-200 shrink-0">
              <button onClick={() => setShowExportPreview(false)} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold">
                <X size={14} /> Close
              </button>
              <button onClick={handleObPrint} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold">
                <Printer size={14} /> Print
              </button>
              <button onClick={handleObExportExcel} disabled={obExporting} className="flex items-center gap-1.5 rounded-xl bg-amber-500 text-stone-900 px-3.5 py-2 text-xs font-bold ml-auto disabled:opacity-50">
                <FileSpreadsheet size={14} /> {obExporting ? "Preparing..." : "Export Excel"}
              </button>
            </div>
          </div>
        </div>
      )}
      {style && obTab !== "newstyle" && obTab !== "templateupload" && obTab !== "savedob" && obTab !== "subentries" && (
        <p className="text-[10px] text-stone-400 px-1 -mt-1">
          Import ஏற்கும் order excel-ல rows இருக்கும் sequence தான் — Heading row-க்கு "Heading" column-ல Y போடு (Section/Machine/SAM/Manpower அந்த row-க்கு தேவையில்லை). Download Template button-ல sample format பாருங்க.
        </p>
      )}

      {!style && obTab !== "newstyle" && obTab !== "templateupload" && obTab !== "savedob" && obTab !== "subentries" && (
        <EmptyState icon={ClipboardList} title="Select or create a style" body="Use the Style panel above to switch to an existing style or create a new one." />
      )}

      {style && (
        <>
          {obTab === "target" && (
          <Section title="Customer Target &amp; Suggested Manpower" icon={Users}>
            <div className="rounded-2xl border border-stone-200 bg-white p-4">
              <Field label="Customer Target Qty / Day" value={style.obTargetQtyPerDay} onChange={(v) => onUpdateStyle({ ...style, obTargetQtyPerDay: v })} suffix="pcs" />
              <div className="mt-2.5 grid grid-cols-2 gap-2.5">
                <Field
                  label="Working Hours / Day"
                  value={style.obWorkingHours ?? ""}
                  onChange={(v) => onUpdateStyle({ ...style, obWorkingHours: v })}
                  suffix="hrs"
                  placeholder="8"
                />
                <Field
                  label="Actual Sewing Manpower"
                  value={style.obActualSewingMp ?? ""}
                  onChange={(v) => onUpdateStyle({ ...style, obActualSewingMp: v })}
                  suffix="heads"
                />
              </div>
              {num(style.obTargetQtyPerDay) > 0 && (
                sewingSam > 0 ? (
                  <div className="mt-2.5 grid grid-cols-3 gap-2.5">
                    <Stat
                      label="Actual Manpower (Sewing)"
                      value={actualSewingMp > 0 ? actualSewingMp : "—"}
                      sub={actualSewingMpManual ? "Entered above" : `Auto: ${totalAllocatedSewingMp} MP allocated in Sewing ops`}
                      accent="text-stone-700"
                    />
                    <Stat
                      label="Suggested Manpower @ 100% Target"
                      value={suggestedMp100 > 0 ? Math.ceil(suggestedMp100) : "—"}
                      sub={`Full target @ 100% efficiency, ${money(sewingSam)} SMV`}
                      accent="text-teal-600"
                    />
                    <Stat
                      label="100% Target Qty"
                      value={achievableQtyAtActualMp > 0 ? Math.floor(achievableQtyAtActualMp) : "—"}
                      sub={actualSewingMp > 0 ? `${actualSewingMp} MP @ 100% eff` : "Enter Actual Sewing Manpower above"}
                      accent={achievableQtyAtActualMp >= num(style.obTargetQtyPerDay) ? "text-emerald-600" : "text-red-600"}
                    />
                  </div>
                ) : (
                  <p className="mt-2 text-[11px] text-stone-400">Add Sewing operations with SAM below to get a suggested manpower.</p>
                )
              )}
            </div>

            <div className="rounded-2xl border border-stone-200 bg-white p-4 mt-2.5">
              <h5 className="text-[11px] font-bold text-stone-500 uppercase tracking-wide mb-2">Manpower Balance (Head Efficiency %)</h5>
              <div className="flex items-end gap-2.5">
                <Field
                  label="Head Efficiency % (target)"
                  value={headEffPct || style.targetEfficiencyPct}
                  onChange={(v) => setHeadEffPct(v)}
                  suffix="%"
                />
                <div className="flex-1">
                  <Stat label="Target Qty / Day" value={style.obTargetQtyPerDay || "—"} sub="from above" />
                </div>
              </div>
              {!(num(style.obTargetQtyPerDay) > 0) ? (
                <p className="mt-2.5 text-xs text-stone-400">Set Customer Target Qty / Day above to see manpower balance per operation.</p>
              ) : mpBalance.rows.filter((r) => r.section === "sewing").length === 0 ? (
                <p className="mt-2.5 text-xs text-stone-400">Add sewing operations with SAM below to see manpower balance.</p>
              ) : (
                <div className="mt-3 space-y-1.5">
                  {mpBalance.rows.filter((r) => r.section === "sewing").map((r) => (
                    <div key={r.id} className="rounded-lg border border-stone-200 bg-stone-50 px-2.5 py-2">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-semibold text-stone-700">{r.name || `Op ${r.slNo}`}</span>
                        <span
                          className={`text-xs font-bold tabular-nums px-1.5 py-0.5 rounded ${
                            r.effPct >= 95 && r.effPct <= 110 ? "bg-emerald-100 text-emerald-700" : r.effPct > 110 ? "bg-red-100 text-red-700" : "bg-amber-100 text-amber-700"
                          }`}
                        >
                          {r.effPct.toFixed(0)}%
                        </span>
                      </div>
                      <div className="mt-1 flex items-center justify-between text-[11px] text-stone-500">
                        <span>Required MP: <b className="text-stone-700 tabular-nums">{r.requiredMp.toFixed(2)}</b></span>
                        <span>Allocated: <b className="text-stone-700 tabular-nums">{r.allocatedMp}</b></span>
                        <span>
                          Balance:{" "}
                          <b className={`tabular-nums ${r.balanceMp < -0.05 ? "text-red-600" : r.balanceMp > 0.05 ? "text-emerald-600" : "text-stone-700"}`}>
                            {r.balanceMp > 0 ? "+" : ""}{r.balanceMp.toFixed(2)}
                          </b>
                        </span>
                      </div>
                    </div>
                  ))}
                  <p className="mt-1 text-[10px] text-stone-400">
                    Required MP = operators needed on that operation to hit Target Qty/Day at the head efficiency % above. Balance = Allocated − Required (negative = short-staffed, positive = spare hands).
                  </p>
                </div>
              )}
            </div>
          </Section>
          )}

          {obTab === "ops" && (
          <>
          <Section title="Operation Bulletin — Sl No / Operation / Section / Machine / SAM / Manpower" icon={ClipboardList}>
            <div className="flex justify-end gap-1.5 mb-2">
              <button
                onClick={() => setOpsView("table")}
                className={`rounded-lg px-3 py-1.5 text-[11px] font-bold border ${opsView === "table" ? "bg-amber-500 border-amber-500 text-stone-900" : "border-stone-200 text-stone-500"}`}
              >
                Table view
              </button>
              <button
                onClick={() => setOpsView("cards")}
                className={`rounded-lg px-3 py-1.5 text-[11px] font-bold border ${opsView === "cards" ? "bg-amber-500 border-amber-500 text-stone-900" : "border-stone-200 text-stone-500"}`}
              >
                Card view
              </button>
            </div>
            {opsView === "table" ? (
              <div className="rounded-2xl border border-stone-200 bg-white overflow-hidden">
                {operations.length === 0 ? (
                  <p className="text-xs text-stone-400 px-3 py-3">No operations added yet. Add cutting, sewing, QC and packing operations below.</p>
                ) : (
                  <div className="overflow-x-auto">
                    <table
                      className="text-xs border-collapse"
                      style={{ width: Object.values(obColWidths).reduce((a, b) => a + b, 0) + 40, tableLayout: "fixed" }}
                    >
                      <colgroup>
                        <col style={{ width: obColWidths.slno }} />
                        <col style={{ width: obColWidths.head }} />
                        <col style={{ width: obColWidths.operation }} />
                        <col style={{ width: obColWidths.section }} />
                        <col style={{ width: obColWidths.machine }} />
                        <col style={{ width: obColWidths.sam }} />
                        <col style={{ width: obColWidths.manpower }} />
                        <col style={{ width: obColWidths.req }} />
                        <col style={{ width: 40 }} />
                      </colgroup>
                      <thead>
                        <tr className="bg-stone-50 text-[10px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
                          <th className="relative px-2 py-2 text-left">#<ObColResizeHandle colKey="slno" /></th>
                          <th className="relative px-2 py-2 text-left">Head<ObColResizeHandle colKey="head" /></th>
                          <th className="relative px-2 py-2 text-left">Operation<ObColResizeHandle colKey="operation" /></th>
                          <th className="relative px-2 py-2 text-left">Section<ObColResizeHandle colKey="section" /></th>
                          <th className="relative px-2 py-2 text-left">Machine<ObColResizeHandle colKey="machine" /></th>
                          <th className="relative px-2 py-2 text-left">SAM<ObColResizeHandle colKey="sam" /></th>
                          <th className="relative px-2 py-2 text-left">Manpower<ObColResizeHandle colKey="manpower" /></th>
                          <th className="relative px-2 py-2 text-left">Req MP / Eff<ObColResizeHandle colKey="req" /></th>
                          <th className="px-2 py-2"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {operations.map((op) => {
                          const mpRow = mpBalance.rows.find((r) => r.id === op.id);
                          return (
                            <React.Fragment key={op.id}>
                              <tr className={`group border-b border-stone-100 align-top ${op.isHeading ? "bg-stone-50" : ""}`}>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={op.slNo}
                                    readOnly
                                    disabled
                                    title="Auto-numbered — updates automatically as rows are added/removed"
                                    placeholder="#"
                                    className="w-9 text-xs text-center rounded-md border border-stone-200 px-1 py-1 outline-none disabled:bg-stone-100 disabled:text-stone-400"
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <div className="flex items-center gap-1">
                                    <input
                                      type="checkbox"
                                      checked={!!op.isHeading}
                                      onChange={(e) => {
                                        updateOperation(op.id, { isHeading: e.target.checked });
                                        if (e.target.checked) saveHeadingToLibrary(op.name);
                                      }}
                                      className="accent-amber-500"
                                    />
                                    {op.isHeading && (
                                      <input
                                        type="checkbox"
                                        checked={false}
                                        disabled={!headingHasSavedGroup(op.name)}
                                        onChange={(e) => { if (e.target.checked) applySubEntriesToHeadingRow(op); }}
                                        title={headingHasSavedGroup(op.name) ? "Load this heading's saved sub-entries" : "No saved group for this heading name yet"}
                                        className="accent-emerald-500 disabled:opacity-30"
                                      />
                                    )}
                                  </div>
                                </td>
                                {op.isHeading ? (
                                  <td colSpan={6} className="px-2 py-1.5">
                                    <div className="flex items-center gap-1.5">
                                      <input
                                        type="text"
                                        list="ob-heading-list"
                                        value={op.name}
                                        onChange={(e) => updateOperation(op.id, { name: e.target.value })}
                                        onBlur={() => saveHeadingToLibrary(op.name)}
                                        onKeyDown={(e) => handleOpEnterKey(e, op.id, "name")}
                                        data-op-row={op.id}
                                        data-op-col="name"
                                        placeholder="Heading (e.g. FRONT PART)"
                                        className="w-full text-xs font-extrabold uppercase tracking-wide rounded-md border border-stone-200 px-2 py-1 outline-none focus:border-amber-500"
                                      />
                                      <button
                                        type="button"
                                        onClick={() => insertOperationAfter(op.id)}
                                        title="Add a row under this heading"
                                        className="shrink-0 opacity-0 group-hover:opacity-100 focus:opacity-100 transition-opacity rounded-md bg-amber-500 text-stone-900 p-1 active:bg-amber-400"
                                      >
                                        <Plus size={12} />
                                      </button>
                                    </div>
                                  </td>
                                ) : (
                                  <>
                                    <td className="px-2 py-1.5">
                                      <input
                                        type="text"
                                        value={op.name}
                                        onChange={(e) => updateOperation(op.id, { name: e.target.value })}
                                        onKeyDown={(e) => handleOpEnterKey(e, op.id, "name")}
                                        data-op-row={op.id}
                                        data-op-col="name"
                                        placeholder="Operation name"
                                        className="w-full text-xs font-semibold rounded-md border border-stone-200 px-2 py-1 outline-none focus:border-amber-500"
                                      />
                                      {op.section !== "qc" && !op.machine && suggestMachine(op.name) && (
                                        <button
                                          type="button"
                                          onClick={() => updateOperation(op.id, { machine: suggestMachine(op.name) })}
                                          className="mt-0.5 text-[9px] font-semibold text-amber-600 active:text-amber-700 whitespace-nowrap"
                                        >
                                          Suggest: {suggestMachine(op.name)}
                                        </button>
                                      )}
                                    </td>
                                    <td className="px-2 py-1.5">
                                      <select
                                        value={op.section}
                                        onChange={(e) => updateOperation(op.id, { section: e.target.value })}
                                        onKeyDown={(e) => handleOpEnterKey(e, op.id, "section")}
                                        data-op-row={op.id}
                                        data-op-col="section"
                                        className="w-full text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500"
                                      >
                                        {OB_SECTIONS.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
                                      </select>
                                    </td>
                                    <td className="px-2 py-1.5">
                                      {op.section === "qc" ? (
                                        <select
                                          value={op.machine}
                                          onChange={(e) => updateOperation(op.id, { machine: e.target.value })}
                                          onKeyDown={(e) => handleOpEnterKey(e, op.id, "machine")}
                                          data-op-row={op.id}
                                          data-op-col="machine"
                                          className="w-full text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500"
                                        >
                                          <option value="">QC Type</option>
                                          <option value="Inline">Inline</option>
                                          <option value="End Line">End Line</option>
                                        </select>
                                      ) : (
                                        <input
                                          type="text"
                                          list={op.section === "packing" ? "ob-packing-list" : op.section === "cutting" ? "ob-cutting-list" : "ob-machine-list"}
                                          value={op.machine}
                                          onChange={(e) => updateOperation(op.id, { machine: e.target.value })}
                                          onKeyDown={(e) => handleOpEnterKey(e, op.id, "machine")}
                                          data-op-row={op.id}
                                          data-op-col="machine"
                                          placeholder={op.section === "packing" ? "Process" : op.section === "cutting" ? "Process" : "Machine"}
                                          className="w-full text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500"
                                        />
                                      )}
                                    </td>
                                    <td className="px-2 py-1.5">
                                      <input
                                        type="number"
                                        value={op.sam}
                                        onChange={(e) => updateOperation(op.id, { sam: e.target.value })}
                                        onKeyDown={(e) => handleOpEnterKey(e, op.id, "sam")}
                                        data-op-row={op.id}
                                        data-op-col="sam"
                                        placeholder="SAM"
                                        className="w-14 text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500 tabular-nums"
                                      />
                                    </td>
                                    <td className="px-2 py-1.5">
                                      {op.section === "sewing" ? (
                                        <div className="flex gap-1">
                                          <input
                                            type="number"
                                            value={op.opCount}
                                            onChange={(e) => updateOpHpQc(op.id, { opCount: e.target.value })}
                                            onKeyDown={(e) => handleOpEnterKey(e, op.id, "opCount")}
                                            data-op-row={op.id}
                                            data-op-col="opCount"
                                            placeholder="Op"
                                            title="Operator"
                                            className="w-11 text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500 tabular-nums"
                                          />
                                          <input
                                            type="number"
                                            value={op.hpCount}
                                            onChange={(e) => updateOpHpQc(op.id, { hpCount: e.target.value })}
                                            onKeyDown={(e) => handleOpEnterKey(e, op.id, "hpCount")}
                                            data-op-row={op.id}
                                            data-op-col="hpCount"
                                            placeholder="Hp"
                                            title="Helper"
                                            className="w-11 text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500 tabular-nums"
                                          />
                                        </div>
                                      ) : (
                                        <input
                                          type="number"
                                          value={op.manpower}
                                          onChange={(e) => updateOperation(op.id, { manpower: e.target.value, opCount: "", hpCount: "" })}
                                          onKeyDown={(e) => handleOpEnterKey(e, op.id, "manpower")}
                                          data-op-row={op.id}
                                          data-op-col="manpower"
                                          placeholder={op.section === "qc" ? "QC" : op.section === "packing" ? (packingRoleLabel(op.machine) || "MP") : "MP"}
                                          className="w-16 text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500 tabular-nums"
                                        />
                                      )}
                                    </td>
                                    <td className="px-2 py-1.5">
                                      {mpRow ? (
                                        <div className="flex flex-col gap-0.5">
                                          <span className="text-stone-500 text-[10px]">MP {mpRow.requiredMp.toFixed(2)}</span>
                                          <span
                                            className={`w-fit font-bold tabular-nums px-1.5 py-0.5 rounded text-[10px] ${
                                              mpRow.effPct >= 95 && mpRow.effPct <= 110 ? "bg-emerald-100 text-emerald-700" : mpRow.effPct > 110 ? "bg-red-100 text-red-700" : "bg-amber-100 text-amber-700"
                                            }`}
                                          >
                                            {mpRow.effPct.toFixed(0)}%
                                          </span>
                                        </div>
                                      ) : (
                                        <span className="text-stone-300">—</span>
                                      )}
                                    </td>
                                  </>
                                )}
                                <td className="px-2 py-1.5">
                                  <button onClick={() => deleteOperation(op.id)} className="text-stone-300 active:text-red-600"><X size={14} /></button>
                                </td>
                              </tr>
                              <tr>
                                <td colSpan={9} className="px-2 py-0.5">
                                  <button
                                    type="button"
                                    onClick={() => insertOperationAfter(op.id)}
                                    className="w-full flex items-center gap-1.5 text-[10px] font-bold text-stone-300 active:text-amber-600 py-0.5 group"
                                  >
                                    <span className="flex-1 border-t border-dashed border-stone-200 group-active:border-amber-300" />
                                    <Plus size={11} /> Insert operation here
                                    <span className="flex-1 border-t border-dashed border-stone-200 group-active:border-amber-300" />
                                  </button>
                                </td>
                              </tr>
                            </React.Fragment>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
                <div className="flex gap-2 p-3">
                  <button onClick={addOperation} className="flex-1 flex items-center justify-center gap-1.5 rounded-xl border-2 border-dashed border-amber-400 text-amber-700 px-3 py-2 text-xs font-bold active:bg-amber-50">
                    <Plus size={14} /> Add Operation
                  </button>
                  <button onClick={addHeading} className="flex-1 flex items-center justify-center gap-1.5 rounded-xl border-2 border-dashed border-stone-300 text-stone-500 px-3 py-2 text-xs font-bold active:bg-stone-50">
                    <Plus size={14} /> Add Heading
                  </button>
                </div>
              </div>
            ) : (
            <div className="rounded-2xl border border-stone-200 bg-white p-3 space-y-2">
              {operations.length === 0 && <p className="text-xs text-stone-400 px-1 py-2">No operations added yet. Add cutting, sewing, QC and packing operations below.</p>}
              {operations.map((op) => (
                <React.Fragment key={op.id}>
                <div className={`rounded-xl border p-2.5 space-y-2 ${op.isHeading ? "border-stone-300 bg-stone-50" : "border-stone-200"}`}>
                  <div className="flex items-center gap-2">
                    <input type="text" value={op.slNo} readOnly disabled title="Auto-numbered — updates automatically as rows are added/removed" placeholder="#" className="w-10 text-xs text-center rounded-md border border-stone-200 px-1 py-1.5 outline-none disabled:bg-stone-100 disabled:text-stone-400" />
                    <label className="flex items-center gap-1 shrink-0 text-[9px] font-bold text-stone-400 uppercase tracking-wide">
                      <input
                        type="checkbox"
                        checked={!!op.isHeading}
                        onChange={(e) => {
                          updateOperation(op.id, { isHeading: e.target.checked });
                          if (e.target.checked) saveHeadingToLibrary(op.name);
                        }}
                        className="accent-amber-500"
                      />
                      Heading
                    </label>
                    {op.isHeading && (
                      <label className={`flex items-center gap-1 shrink-0 text-[9px] font-bold uppercase tracking-wide ${headingHasSavedGroup(op.name) ? "text-emerald-600" : "text-stone-300"}`}>
                        <input
                          type="checkbox"
                          checked={false}
                          disabled={!headingHasSavedGroup(op.name)}
                          onChange={(e) => { if (e.target.checked) applySubEntriesToHeadingRow(op); }}
                          title={headingHasSavedGroup(op.name) ? "Load this heading's saved sub-entries" : "No saved group for this heading name yet"}
                          className="accent-emerald-500 disabled:opacity-30"
                        />
                        Load
                      </label>
                    )}
                    <input
                      type="text"
                      list={op.isHeading ? "ob-heading-list" : undefined}
                      value={op.name}
                      onChange={(e) => updateOperation(op.id, { name: e.target.value })}
                      onBlur={() => { if (op.isHeading) saveHeadingToLibrary(op.name); }}
                      placeholder={op.isHeading ? "Heading (e.g. FRONT PART)" : "Operation name (e.g. Shoulder attach)"}
                      className={`flex-1 text-xs rounded-md border border-stone-200 px-2 py-1.5 outline-none focus:border-amber-500 ${op.isHeading ? "font-extrabold uppercase tracking-wide" : "font-semibold"}`}
                    />
                    <button onClick={() => deleteOperation(op.id)} className="text-stone-300 active:text-red-600 shrink-0"><X size={15} /></button>
                  </div>
                  {!op.isHeading && (
                    <>
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-1.5">
                        <select value={op.section} onChange={(e) => updateOperation(op.id, { section: e.target.value })} className="text-[11px] rounded-md border border-stone-200 px-1.5 py-1.5 outline-none focus:border-amber-500">
                          {OB_SECTIONS.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
                        </select>
                        {op.section === "qc" ? (
                          <select value={op.machine} onChange={(e) => updateOperation(op.id, { machine: e.target.value })} className="text-[11px] rounded-md border border-stone-200 px-1.5 py-1.5 outline-none focus:border-amber-500">
                            <option value="">QC Type</option>
                            <option value="Inline">Inline</option>
                            <option value="End Line">End Line</option>
                          </select>
                        ) : (
                          <input
                            type="text"
                            list={op.section === "packing" ? "ob-packing-list" : op.section === "cutting" ? "ob-cutting-list" : "ob-machine-list"}
                            value={op.machine}
                            onChange={(e) => updateOperation(op.id, { machine: e.target.value })}
                            placeholder={op.section === "packing" ? "Packing process (Ironing, Metal Detection...)" : op.section === "cutting" ? "Cutting process (Layering, Cutter...)" : "Machine (SNLS, O/L...)"}
                            className="text-[11px] rounded-md border border-stone-200 px-1.5 py-1.5 outline-none focus:border-amber-500"
                          />
                        )}
                        <input type="number" value={op.sam} onChange={(e) => updateOperation(op.id, { sam: e.target.value })} placeholder="SAM" className="text-[11px] rounded-md border border-stone-200 px-1.5 py-1.5 outline-none focus:border-amber-500 tabular-nums" />
                        {op.section === "sewing" ? (
                          <div className="flex gap-1">
                            <input
                              type="number"
                              value={op.opCount}
                              onChange={(e) => updateOpHpQc(op.id, { opCount: e.target.value })}
                              placeholder="Operator"
                              title="Operator"
                              className="w-1/2 text-[11px] rounded-md border border-stone-200 px-1.5 py-1.5 outline-none focus:border-amber-500 tabular-nums"
                            />
                            <input
                              type="number"
                              value={op.hpCount}
                              onChange={(e) => updateOpHpQc(op.id, { hpCount: e.target.value })}
                              placeholder="Helper"
                              title="Helper"
                              className="w-1/2 text-[11px] rounded-md border border-stone-200 px-1.5 py-1.5 outline-none focus:border-amber-500 tabular-nums"
                            />
                          </div>
                        ) : (
                          <input
                            type="number"
                            value={op.manpower}
                            onChange={(e) => updateOperation(op.id, { manpower: e.target.value, opCount: "", hpCount: "" })}
                            placeholder={op.section === "qc" ? "QC" : op.section === "packing" ? (packingRoleLabel(op.machine) || "Manpower") : "Manpower"}
                            className="text-[11px] rounded-md border border-stone-200 px-1.5 py-1.5 outline-none focus:border-amber-500 tabular-nums"
                          />
                        )}
                      </div>
                    </>
                  )}
                  {!op.isHeading && op.section !== "qc" && !op.machine && suggestMachine(op.name) && (
                    <button type="button" onClick={() => updateOperation(op.id, { machine: suggestMachine(op.name) })} className="text-[10px] font-semibold text-amber-600 active:text-amber-700">
                      Suggested machine: {suggestMachine(op.name)} — tap to use (or type your own above)
                    </button>
                  )}
                  {!op.isHeading && (() => {
                    const mpRow = mpBalance.rows.find((r) => r.id === op.id);
                    if (!mpRow) return null;
                    return (
                      <div className="flex items-center justify-between rounded-md bg-stone-50 border border-stone-200 px-2 py-1.5 text-[11px]">
                        <span className="text-stone-500">Required MP: <b className="text-stone-700 tabular-nums">{mpRow.requiredMp.toFixed(2)}</b></span>
                        <span
                          className={`font-bold tabular-nums px-1.5 py-0.5 rounded ${
                            mpRow.effPct >= 95 && mpRow.effPct <= 110 ? "bg-emerald-100 text-emerald-700" : mpRow.effPct > 110 ? "bg-red-100 text-red-700" : "bg-amber-100 text-amber-700"
                          }`}
                        >
                          {mpRow.effPct.toFixed(0)}%
                        </span>
                      </div>
                    );
                  })()}
                </div>
                <button
                  type="button"
                  onClick={() => insertOperationAfter(op.id)}
                  className="w-full flex items-center gap-1.5 text-[10px] font-bold text-stone-300 active:text-amber-600 py-0.5 group"
                >
                  <span className="flex-1 border-t border-dashed border-stone-200 group-active:border-amber-300" />
                  <Plus size={11} /> Insert operation here
                  <span className="flex-1 border-t border-dashed border-stone-200 group-active:border-amber-300" />
                </button>
                </React.Fragment>
              ))}
              <div className="flex gap-2">
                <button onClick={addOperation} className="flex-1 flex items-center justify-center gap-1.5 rounded-xl border-2 border-dashed border-amber-400 text-amber-700 px-3 py-2 text-xs font-bold active:bg-amber-50">
                  <Plus size={14} /> Add Operation
                </button>
                <button onClick={addHeading} className="flex-1 flex items-center justify-center gap-1.5 rounded-xl border-2 border-dashed border-stone-300 text-stone-500 px-3 py-2 text-xs font-bold active:bg-stone-50">
                  <Plus size={14} /> Add Heading
                </button>
              </div>
            </div>
            )}
          </Section>

          <Section title="Section-wise Usage Summary" icon={GitBranch}>
            <div className="rounded-2xl border border-stone-200 bg-white p-3 space-y-2">
              <div className="grid grid-cols-2 gap-2">
                {OB_SECTIONS.map((s) => (
                  <div key={s.key} className="rounded-lg bg-stone-50 border border-stone-200 px-2.5 py-2">
                    <p className="text-[10px] font-bold text-stone-500 uppercase tracking-wide">{s.label}</p>
                    <p className="text-sm font-bold text-stone-700 tabular-nums">{summary.sections[s.key].count} operation(s)</p>
                    <p className="text-[11px] text-stone-400">SMV: <b className="text-stone-600">{money(summary.sections[s.key].smv)}</b> · Manpower: <b className="text-stone-600">{summary.sections[s.key].manpower}</b></p>
                  </div>
                ))}
              </div>
              {filledSections.length > 0 && (
                <div className="rounded-lg bg-amber-50 border border-amber-200 px-2.5 py-2 flex items-center justify-between">
                  <span className="text-xs font-semibold text-amber-700">
                    {filledSections.length > 1 ? filledSections.map((s) => s.label).join(" + ") : `${filledSections[0].label} SMV`}
                  </span>
                  <span className="text-sm font-bold text-amber-700 tabular-nums">{money(filledSectionsSmv)}</span>
                </div>
              )}
            </div>
          </Section>

          {(machineUsage.length > 0 || helperSummary.total > 0) && (
            <Section title="Machine-wise Usage" icon={Settings}>
              <div className="rounded-2xl border border-stone-200 bg-white p-3 space-y-1.5">
                <div className="flex items-center justify-between rounded-lg bg-amber-50 border border-amber-200 px-2.5 py-1.5 mb-1">
                  <span className="text-xs font-semibold text-amber-700">Total Machines / Manpower</span>
                  <span className="text-sm font-bold text-amber-700 tabular-nums">{machineUsage.reduce((a, m) => a + m.manpower, 0)}</span>
                </div>
                {machineUsage.map((m) => (
                  <div key={m.machine} className="flex items-center justify-between rounded-lg bg-stone-50 border border-stone-200 px-2.5 py-1.5">
                    <span className="text-xs font-semibold text-stone-600">{m.machine}</span>
                    <span className="text-[11px] text-stone-400">{m.opCount} operation(s) · <b className="text-stone-600">{m.manpower}</b> manpower/machines</span>
                  </div>
                ))}
                {helperSummary.total > 0 && (
                  <div className="grid grid-cols-2 gap-1.5 pt-1">
                    <div className="rounded-lg bg-sky-50 border border-sky-200 px-2.5 py-1.5">
                      <p className="text-[10px] font-semibold text-sky-600 uppercase tracking-wide">Total Operators (OP)</p>
                      <p className="text-sm font-bold text-sky-700 tabular-nums">{helperSummary.totalOp}</p>
                    </div>
                    <div className="rounded-lg bg-violet-50 border border-violet-200 px-2.5 py-1.5">
                      <p className="text-[10px] font-semibold text-violet-600 uppercase tracking-wide">Total Helpers (HP)</p>
                      <p className="text-sm font-bold text-violet-700 tabular-nums">{helperSummary.totalHp}</p>
                    </div>
                  </div>
                )}
              </div>
            </Section>
          )}
          </>
          )}

          {obTab === "subentries" && (
          <>
          <Section title="Add Group — Heading + all its rows in one table (like an Operation Breakdown sheet)" icon={Boxes}>
            <div className="rounded-2xl border border-stone-200 bg-white p-3 space-y-2.5">
              <p className="text-[11px] text-stone-400">Type a Heading (e.g. THUMB HOLE) and its full row of operations at once, save it as one ready group, then insert the whole group into any style with a single tap.</p>
              <div className="flex flex-wrap gap-2">
                <button onClick={handleSubExportExcel} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold">
                  <FileSpreadsheet size={14} /> Export Excel
                </button>
                <button onClick={handleSubDownloadTemplate} className="flex items-center gap-1.5 rounded-xl border border-stone-200 text-stone-600 px-3.5 py-2 text-xs font-bold">
                  <Download size={14} /> Sample Template
                </button>
                <button onClick={() => subFileInputRef.current?.click()} className="flex items-center gap-1.5 rounded-xl border-2 border-dashed border-amber-400 text-amber-700 px-3.5 py-2 text-xs font-bold active:bg-amber-50">
                  <Upload size={14} /> Import Excel
                </button>
                <input ref={subFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleSubImportExcel} className="hidden" />
              </div>
              <input
                type="text"
                list="ob-heading-list"
                value={bulkHeading}
                onChange={(e) => setBulkHeading(e.target.value)}
                placeholder="Heading (e.g. THUMB HOLE)"
                className="w-full text-sm font-extrabold uppercase tracking-wide text-center rounded-xl border border-stone-200 bg-amber-50 px-3 py-2.5 outline-none focus:border-amber-500"
              />
              <PhotoRefUpload value={bulkPhotoRef} onChange={setBulkPhotoRef} label="Group Photo Reference" />
              <div className="overflow-x-auto rounded-xl border border-stone-200">
                <table className="w-full min-w-[640px] text-xs border-collapse">
                  <thead>
                    <tr className="bg-stone-50 text-[10px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
                      <th className="px-2 py-2 text-left w-10">#</th>
                      <th className="px-2 py-2 text-left min-w-[160px]">Operation</th>
                      <th className="px-2 py-2 text-left w-24">Section</th>
                      <th className="px-2 py-2 text-left w-28">Machine</th>
                      <th className="px-2 py-2 text-left w-16">SAM</th>
                      <th className="px-2 py-2 text-left w-12">OP</th>
                      <th className="px-2 py-2 text-left w-12">HP</th>
                      <th className="px-2 py-2 w-8"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {bulkRows.map((row, i) => (
                      <tr key={row.id} className="border-b border-stone-100 align-top">
                        <td className="px-2 py-1.5 text-stone-400 font-semibold">{i + 1}</td>
                        <td className="px-2 py-1.5">
                          <input
                            type="text"
                            value={row.name}
                            onChange={(e) => updateBulkRow(row.id, { name: e.target.value })}
                            placeholder="Operation name"
                            className="w-full text-xs font-semibold rounded-md border border-stone-200 px-1.5 py-1 outline-none focus:border-amber-500"
                          />
                          {row.section !== "qc" && !row.machine && suggestMachine(row.name) && (
                            <button
                              type="button"
                              onClick={() => updateBulkRow(row.id, { machine: suggestMachine(row.name) })}
                              className="mt-0.5 text-[9px] font-semibold text-amber-600 active:text-amber-700 whitespace-nowrap"
                            >
                              Suggest: {suggestMachine(row.name)}
                            </button>
                          )}
                        </td>
                        <td className="px-2 py-1.5">
                          <select
                            value={row.section}
                            onChange={(e) => updateBulkRow(row.id, { section: e.target.value })}
                            className="w-full text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500"
                          >
                            {OB_SECTIONS.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
                          </select>
                        </td>
                        <td className="px-2 py-1.5">
                          {row.section === "qc" ? (
                            <select
                              value={row.machine}
                              onChange={(e) => updateBulkRow(row.id, { machine: e.target.value })}
                              className="w-full text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500"
                            >
                              <option value="">QC Type</option>
                              <option value="Inline">Inline</option>
                              <option value="End Line">End Line</option>
                            </select>
                          ) : (
                            <input
                              type="text"
                              list={row.section === "packing" ? "ob-packing-list" : row.section === "cutting" ? "ob-cutting-list" : "ob-machine-list"}
                              value={row.machine}
                              onChange={(e) => updateBulkRow(row.id, { machine: e.target.value })}
                              placeholder={row.section === "packing" || row.section === "cutting" ? "Process" : "Machine"}
                              className="w-full text-[11px] rounded-md border border-stone-200 px-1.5 py-1 outline-none focus:border-amber-500"
                            />
                          )}
                        </td>
                        <td className="px-2 py-1.5">
                          <input
                            type="number"
                            value={row.sam}
                            onChange={(e) => updateBulkRow(row.id, { sam: e.target.value })}
                            placeholder="SAM"
                            className="w-14 text-[11px] rounded-md border border-stone-200 px-1.5 py-1 outline-none focus:border-amber-500 tabular-nums"
                          />
                        </td>
                        <td className="px-2 py-1.5">
                          <input
                            type="number"
                            value={row.opCount}
                            onChange={(e) => updateBulkRow(row.id, { opCount: e.target.value })}
                            placeholder="Op"
                            className="w-11 text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500 tabular-nums"
                          />
                        </td>
                        <td className="px-2 py-1.5">
                          <input
                            type="number"
                            value={row.hpCount}
                            onChange={(e) => updateBulkRow(row.id, { hpCount: e.target.value })}
                            placeholder="Hp"
                            className="w-11 text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500 tabular-nums"
                          />
                        </td>
                        <td className="px-2 py-1.5">
                          <button onClick={() => removeBulkRow(row.id)} className="text-stone-300 active:text-red-600"><X size={14} /></button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex gap-2">
                <button type="button" onClick={addBulkRow} className="flex-1 flex items-center justify-center gap-1.5 rounded-xl border-2 border-dashed border-amber-400 text-amber-700 px-3 py-2 text-xs font-bold active:bg-amber-50">
                  <Plus size={14} /> Add Row
                </button>
                <button
                  onClick={handleSaveBulkRows}
                  disabled={!bulkRows.some((r) => r.name.trim())}
                  className="flex-1 flex items-center justify-center gap-1.5 rounded-xl bg-amber-500 disabled:opacity-40 text-stone-900 font-bold px-3 py-2 text-xs active:bg-amber-400"
                >
                  <Check size={14} /> Save Group
                </button>
              </div>
            </div>
          </Section>

          {subEntryGroups.length > 0 && (
          <Section title={`Saved Groups (${subEntryGroups.length})`} icon={Layers}>
            <div className="rounded-2xl border border-stone-200 bg-white p-3 space-y-2">
              <p className="text-[11px] text-stone-400">Each group's rows are saved together under its Heading — tap "Select" to filter and pick the ones you want; each pick drops in at the top of that heading, pushing the existing rows down, and the list stays open for more picks.</p>
              {subEntryGroups.map((g) => {
                const groupKey = g.heading.toLowerCase();
                const isOpen = openGroupKey === groupKey;
                const isEditOpen = editGroupKey === groupKey;
                const filteredRows = g.rows.filter(
                  (r) => !groupFilterText.trim() || (r.name || "").toLowerCase().includes(groupFilterText.trim().toLowerCase())
                );
                return (
                  <div key={groupKey} className="rounded-xl border border-stone-200 bg-stone-50 px-3 py-2.5 space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <div className="min-w-0 flex items-center gap-2">
                        {getGroupPhoto(g.heading) && (
                          <button
                            type="button"
                            onClick={() => { setEditGroupKey(isEditOpen ? "" : groupKey); setOpenGroupKey(""); }}
                            title="Tap to preview / change photo"
                            className="shrink-0"
                          >
                            <img src={getGroupPhoto(g.heading)} alt={g.heading} className="w-16 h-16 rounded-lg object-cover border border-stone-200" />
                          </button>
                        )}
                        <div className="min-w-0">
                          <p className="text-xs font-extrabold uppercase tracking-wide text-amber-700 truncate">{g.heading}</p>
                          <p className="text-[10px] text-stone-400">{g.rows.length} operation(s)</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-1.5 shrink-0">
                        <button
                          type="button"
                          onClick={() => { setEditGroupKey(isEditOpen ? "" : groupKey); setOpenGroupKey(""); }}
                          className="flex items-center gap-1 rounded-md border border-dashed border-stone-300 text-stone-500 font-bold px-2.5 py-1.5 text-[11px] active:bg-stone-100 whitespace-nowrap"
                        >
                          <Pencil size={12} /> {isEditOpen ? "Close" : "Edit"}
                        </button>
                        <button
                          type="button"
                          onClick={() => insertSubEntryGroup(g.heading, g.rows)}
                          disabled={!style}
                          title="Insert every operation in this group at once"
                          className="flex items-center gap-1 rounded-md border border-dashed border-emerald-400 text-emerald-600 disabled:opacity-40 disabled:border-stone-200 disabled:text-stone-300 font-bold px-2.5 py-1.5 text-[11px] active:bg-emerald-100 whitespace-nowrap"
                        >
                          <Layers size={12} /> Insert
                        </button>
                        <button
                          type="button"
                          onClick={() => { setOpenGroupKey(isOpen ? "" : groupKey); setGroupFilterText(""); setEditGroupKey(""); }}
                          disabled={!style}
                          className="flex items-center gap-1 rounded-md border border-dashed border-amber-400 text-amber-600 disabled:opacity-40 disabled:border-stone-200 disabled:text-stone-300 font-bold px-2.5 py-1.5 text-[11px] active:bg-amber-100 whitespace-nowrap"
                        >
                          {isOpen ? <X size={12} /> : <Plus size={12} />} {isOpen ? "Close" : "Select"}
                        </button>
                        <button onClick={() => deleteSubEntryGroup(g.heading)} className="text-stone-300 active:text-red-600 shrink-0"><X size={14} /></button>
                      </div>
                    </div>
                    {isEditOpen && (
                      <div className="rounded-lg border border-stone-200 bg-white p-2 space-y-2.5">
                        <PhotoRefUpload
                          value={getGroupPhoto(g.heading)}
                          onChange={(dataUri) => saveGroupPhoto(g.heading, dataUri)}
                          label="Group Photo Reference"
                        />
                        <div className="overflow-x-auto">
                          <table className="w-full min-w-[760px] text-xs border-collapse">
                            <thead>
                              <tr className="bg-stone-50 text-[10px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
                                <th className="px-2 py-2 text-left w-28">Heading</th>
                                <th className="px-2 py-2 text-left min-w-[170px]">Operation</th>
                                <th className="px-2 py-2 text-left w-24">Section</th>
                                <th className="px-2 py-2 text-left w-28">Machine</th>
                                <th className="px-2 py-2 text-left w-14">SAM</th>
                                <th className="px-2 py-2 text-left w-24">MP</th>
                                <th className="px-2 py-2 w-10"></th>
                              </tr>
                            </thead>
                            <tbody>
                              {g.rows.map((entry) => (
                                <tr key={entry.id} className="border-b border-stone-100 align-top">
                                  <td className="px-2 py-1.5">
                                    <input
                                      type="text"
                                      list="ob-heading-list"
                                      value={entry.heading}
                                      onChange={(e) => updateSubEntry(entry.id, { heading: e.target.value })}
                                      placeholder="(none)"
                                      className="w-full text-[11px] font-bold text-amber-600 uppercase tracking-wide rounded-md border border-stone-200 px-1.5 py-1 outline-none focus:border-amber-500"
                                    />
                                  </td>
                                  <td className="px-2 py-1.5">
                                    <input
                                      type="text"
                                      value={entry.name}
                                      onChange={(e) => updateSubEntry(entry.id, { name: e.target.value })}
                                      placeholder="Operation name"
                                      className="w-full text-xs font-semibold rounded-md border border-stone-200 px-1.5 py-1 outline-none focus:border-amber-500"
                                    />
                                  </td>
                                  <td className="px-2 py-1.5">
                                    <select
                                      value={entry.section}
                                      onChange={(e) => updateSubEntry(entry.id, { section: e.target.value })}
                                      className="w-full text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500"
                                    >
                                      {OB_SECTIONS.map((s) => <option key={s.key} value={s.key}>{s.label}</option>)}
                                    </select>
                                  </td>
                                  <td className="px-2 py-1.5">
                                    {entry.section === "qc" ? (
                                      <select
                                        value={entry.machine}
                                        onChange={(e) => updateSubEntry(entry.id, { machine: e.target.value })}
                                        className="w-full text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500"
                                      >
                                        <option value="">QC Type</option>
                                        <option value="Inline">Inline</option>
                                        <option value="End Line">End Line</option>
                                      </select>
                                    ) : (
                                      <>
                                        <input
                                          type="text"
                                          list={entry.section === "packing" ? "ob-packing-list" : entry.section === "cutting" ? "ob-cutting-list" : "ob-machine-list"}
                                          value={entry.machine}
                                          onChange={(e) => updateSubEntry(entry.id, { machine: e.target.value })}
                                          placeholder={entry.section === "packing" || entry.section === "cutting" ? "Process" : "Machine"}
                                          className="w-full text-[11px] rounded-md border border-stone-200 px-1.5 py-1 outline-none focus:border-amber-500"
                                        />
                                        {!entry.machine && suggestMachine(entry.name) && (
                                          <button
                                            type="button"
                                            onClick={() => updateSubEntry(entry.id, { machine: suggestMachine(entry.name) })}
                                            className="mt-0.5 text-[9px] font-semibold text-amber-600 active:text-amber-700 whitespace-nowrap"
                                          >
                                            Suggest: {suggestMachine(entry.name)}
                                          </button>
                                        )}
                                      </>
                                    )}
                                  </td>
                                  <td className="px-2 py-1.5">
                                    <input
                                      type="number"
                                      value={entry.sam}
                                      onChange={(e) => updateSubEntry(entry.id, { sam: e.target.value })}
                                      placeholder="SAM"
                                      className="w-14 text-[11px] rounded-md border border-stone-200 px-1.5 py-1 outline-none focus:border-amber-500 tabular-nums"
                                    />
                                  </td>
                                  <td className="px-2 py-1.5">
                                    {entry.section === "sewing" ? (
                                      <div className="flex gap-1">
                                        <input
                                          type="number"
                                          value={entry.opCount || ""}
                                          onChange={(e) => updateSubEntryOpHpQc(entry.id, { opCount: e.target.value })}
                                          placeholder="Op"
                                          title="Operator"
                                          className="w-11 text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500 tabular-nums"
                                        />
                                        <input
                                          type="number"
                                          value={entry.hpCount || ""}
                                          onChange={(e) => updateSubEntryOpHpQc(entry.id, { hpCount: e.target.value })}
                                          placeholder="Hp"
                                          title="Helper"
                                          className="w-11 text-[11px] rounded-md border border-stone-200 px-1 py-1 outline-none focus:border-amber-500 tabular-nums"
                                        />
                                      </div>
                                    ) : (
                                      <input
                                        type="number"
                                        value={entry.manpower}
                                        onChange={(e) => updateSubEntry(entry.id, { manpower: e.target.value, opCount: "", hpCount: "" })}
                                        placeholder={entry.section === "qc" ? "QC" : entry.section === "packing" ? (packingRoleLabel(entry.machine) || "MP") : "MP"}
                                        className="w-14 text-[11px] rounded-md border border-stone-200 px-1.5 py-1 outline-none focus:border-amber-500 tabular-nums"
                                      />
                                    )}
                                  </td>
                                  <td className="px-2 py-1.5">
                                    <button onClick={() => deleteSubEntry(entry.id)} className="text-stone-300 active:text-red-600 shrink-0"><X size={14} /></button>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                        <button
                          type="button"
                          onClick={() => addBlankSubEntryToHeading(g.heading)}
                          className="w-full flex items-center justify-center gap-1.5 rounded-lg border-2 border-dashed border-amber-400 text-amber-700 px-3 py-1.5 text-[11px] font-bold active:bg-amber-50"
                        >
                          <Plus size={12} /> Add Operation to this Group
                        </button>
                      </div>
                    )}
                    {isOpen && (
                      <div className="rounded-lg border border-amber-200 bg-white p-2 space-y-1.5">
                        <div className="relative">
                          <Search size={12} className="absolute left-2 top-1/2 -translate-y-1/2 text-stone-300" />
                          <input
                            type="text"
                            value={groupFilterText}
                            onChange={(e) => setGroupFilterText(e.target.value)}
                            placeholder="Filter operations..."
                            autoFocus
                            className="w-full text-[11px] rounded-md border border-stone-200 pl-6 pr-2 py-1.5 outline-none focus:border-amber-500"
                          />
                        </div>
                        <div className="max-h-48 overflow-y-auto space-y-1">
                          {filteredRows.length === 0 ? (
                            <p className="text-[10px] text-stone-400 px-1 py-1.5">No match.</p>
                          ) : (
                            filteredRows.map((entry) => {
                              const added = isEntryInHeading(g.heading, entry.name);
                              return (
                                <button
                                  key={entry.id}
                                  type="button"
                                  onClick={() => insertSubEntryTop(g.heading, entry)}
                                  className={`w-full flex items-center justify-between gap-2 rounded-md px-2 py-1.5 text-[11px] text-left ${
                                    added ? "bg-emerald-50 text-emerald-700" : "bg-stone-50 text-stone-700 active:bg-amber-100"
                                  }`}
                                >
                                  <span className="truncate">{entry.name || "(unnamed)"}</span>
                                  {added ? <Check size={12} className="shrink-0" /> : <Plus size={12} className="shrink-0 text-amber-600" />}
                                </button>
                              );
                            })
                          )}
                        </div>
                        <p className="text-[10px] text-stone-400 px-1">Tap an operation to add it — already-added ones show a ✓ but you can tap again to add another copy.</p>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </Section>
          )}

          </>
          )}

          {obTab === "balance" && (
          <>
          <Section title="Section-wise SMV" icon={Gauge}>
            <div className="rounded-2xl border border-stone-200 bg-white p-4 grid grid-cols-2 gap-2.5">
              <Stat label="Cutting SMV" value={money(summary.sections.cutting.smv)} accent="text-stone-700" />
              <Stat label="Sewing SMV" value={money(summary.sections.sewing.smv)} accent="text-stone-700" />
              <Stat label="QC SMV" value={money(summary.sections.qc.smv)} accent="text-stone-700" />
              <Stat label="Packing SMV" value={money(summary.sections.packing.smv)} accent="text-stone-700" />
              <Stat label="SMV (Sew to QC)" value={money(summary.sewToQcSmv)} accent="text-amber-600" />
              <Stat label="SMV (Sew to Pack)" value={money(summary.sewToPackSmv)} accent="text-amber-600" />
              <Stat label="Total SMV (incl. Cutting)" value={money(summary.totalSmv)} accent="text-emerald-600" />
            </div>
            <button onClick={syncToStyleSMV} className="mt-2.5 w-full rounded-xl bg-amber-500 text-stone-900 font-bold py-2.5 text-sm active:bg-amber-400">
              Sync to Style SMV (Cutting / Sewing / Packing)
            </button>
            {savedMsg && <p className="mt-1.5 text-center text-xs text-emerald-600 font-semibold">{savedMsg}</p>}
          </Section>

          <Section title="CM Calculation" icon={TrendingUp}>
            <div className="rounded-2xl border border-stone-200 bg-white p-4 space-y-3">
              <div className="grid grid-cols-2 gap-2.5">
                <label className="flex flex-col gap-1">
                  <span className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">Currency</span>
                  <select
                    value={cmCurrency}
                    onChange={(e) => onUpdateStyle({ ...style, obCmCurrency: e.target.value })}
                    className="rounded-xl border border-stone-200 bg-white px-3 py-2.5 text-sm outline-none focus:border-amber-500"
                  >
                    {CM_CURRENCIES.map((c) => (
                      <option key={c.code} value={c.code}>{c.code} ({c.symbol})</option>
                    ))}
                  </select>
                </label>
                <label className="flex flex-col gap-1">
                  <span className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">CM Basis</span>
                  <div className="flex rounded-xl border border-stone-200 overflow-hidden">
                    <button
                      type="button"
                      onClick={() => onUpdateStyle({ ...style, obCmBasis: "dzn" })}
                      className={`flex-1 py-2.5 text-xs font-bold ${cmBasis === "dzn" ? "bg-amber-500 text-stone-900" : "bg-white text-stone-500"}`}
                    >
                      Per DZN
                    </button>
                    <button
                      type="button"
                      onClick={() => onUpdateStyle({ ...style, obCmBasis: "pcs" })}
                      className={`flex-1 py-2.5 text-xs font-bold ${cmBasis === "pcs" ? "bg-amber-500 text-stone-900" : "bg-white text-stone-500"}`}
                    >
                      Per PCS
                    </button>
                  </div>
                </label>
              </div>
              <Field label={`CM / Minute (${cmSymbol})`} value={style.obCmPerMinute || ""} onChange={(v) => onUpdateStyle({ ...style, obCmPerMinute: v })} placeholder="e.g. 0.045" />
              <Field
                label="Avg Line Efficiency %"
                value={style.obCmAvgEffPct || ""}
                onChange={(v) => onUpdateStyle({ ...style, obCmAvgEffPct: v })}
                placeholder={avgSewingEffPct ? `e.g. ${money(avgSewingEffPct)} (from sewing ops)` : "e.g. 65"}
                suffix="%"
              />
              {cmPerMinute > 0 ? (
                <div className="grid grid-cols-2 gap-2.5">
                  <Stat
                    label={`CM / ${cmBasis === "pcs" ? "PCS" : "DZN"} ${cmSymbol}`}
                    value={`${cmSymbol}${money(cmPerDzn)}`}
                    accent="text-amber-600"
                    sub={`SMV ÷ Avg Eff% × CM/min${cmBasis === "pcs" ? "" : " × 12"}`}
                  />
                  <Stat label={`CM / Machine ${cmSymbol}`} value={`${cmSymbol}${money(cmPerMachine)}`} sub={`CM/${cmBasis === "pcs" ? "PCS" : "DZN"} ÷ Total Manpower`} />
                </div>
              ) : (
                <p className="text-xs text-stone-400">CM / Minute ({cmSymbol}) rate போடுங்க — CM/{cmBasis === "pcs" ? "PCS" : "DZN"}, CM/Machine auto calculate ஆகும்.</p>
              )}
            </div>
          </Section>

          <Section title="Line Balancing (Sewing operations)" icon={GitBranch}>
            <div className="rounded-2xl border border-stone-200 bg-white p-4">
              {balancing.rows.length === 0 ? (
                <p className="text-xs text-stone-400">Add sewing operations with SAM & Manpower above to see line balancing.</p>
              ) : (
                <>
                  <div className="grid grid-cols-3 gap-2.5 mb-3">
                    <Stat label="Bottleneck / hr" value={`${money(balancing.lcl)} pcs`} accent="text-red-600" sub={balancing.bottleneck?.name || ""} />
                    <Stat label="Avg Capacity / hr" value={`${money(balancing.avgCapacity)} pcs`} accent="text-stone-700" />
                    <Stat label="Highest (UCL) / hr" value={`${money(balancing.ucl)} pcs`} accent="text-emerald-600" />
                  </div>
                  <div className="h-56">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={chartData} margin={{ top: 4, right: 8, left: -20, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} />
                        <XAxis dataKey="name" tick={{ fontSize: 9 }} interval={0} angle={-30} textAnchor="end" height={60} />
                        <YAxis tick={{ fontSize: 10 }} />
                        <Tooltip />
                        <Bar dataKey="capacity" radius={[4, 4, 0, 0]}>
                          {chartData.map((row, i) => (
                            <Cell key={i} fill={row.capacity <= balancing.lcl + 0.001 ? "#dc2626" : "#0f766e"} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                  <p className="mt-2 text-[11px] text-stone-400">
                    The line's hourly output can't exceed its bottleneck operation ({balancing.bottleneck?.name || "—"}, red bar) — balance manpower or split that operation to raise it toward the average.
                  </p>
                </>
              )}
            </div>
          </Section>
          </>
          )}

          {obTab === "rampup" && (
          <Section title="Order Ramp-up Plan (Day-wise Target)" icon={TrendingUp}>
            <div className="rounded-2xl border border-stone-200 bg-white p-4">
              <div className="grid grid-cols-2 gap-2.5 mb-3">
                <Stat label="Order Qty" value={style.orderQty || "—"} sub="from Order Book" />
                <Field label="No. of Lines" value={style.obPlannedLines || ""} onChange={(v) => onUpdateStyle({ ...style, obPlannedLines: v })} />
              </div>
              <div className="grid grid-cols-2 gap-2.5 mb-3">
                <Field
                  label="Target Qty / Day / Line"
                  value={style.obRampPerLineTarget || ""}
                  onChange={(v) => onUpdateStyle({ ...style, obRampPerLineTarget: v })}
                  placeholder={linePerDayTarget > 0 ? `e.g. ${money(linePerDayTarget)}` : "e.g. 1200"}
                  suffix="pcs"
                />
                <Stat
                  label="Total Line Target / Day"
                  value={noOfLines > 0 ? money(rampTotalTarget) : "—"}
                  sub={noOfLines > 0 ? `${money(rampPerLineTarget)} × ${noOfLines} lines` : "Enter No. of Lines"}
                  accent="text-amber-600"
                />
              </div>
              <div className="grid grid-cols-2 gap-2.5 mb-3">
                <Stat label="Total Manpower" value={rampUpTotalMp > 0 ? money(rampUpTotalMp) : "—"} sub={noOfLines > 0 ? `${money(totalObManpower)} × ${noOfLines} lines` : "Cutting+Sewing+QC+Packing MP entered"} />
                <Stat label="Avg Efficiency %" value={rampUpAvgEffPct > 0 ? `${money(rampUpAvgEffPct)}%` : "—"} sub={`avg of daily Eff% @ ${money(rampUpSam)} SAM`} accent="text-emerald-600" />
              </div>

              <h5 className="text-[11px] font-bold text-stone-500 uppercase tracking-wide mb-2">Ramp-up Days (manual entry, up to peak — enter QTY PER LINE)</h5>
              <div className="space-y-1.5 mb-2.5">
                {(style.obRampUpDays || []).map((r, i) => (
                  <div key={r.id} className="flex items-center gap-1.5">
                    <span className="w-14 text-xs font-semibold text-stone-500">Day {i + 1}</span>
                    <input type="number" value={r.qty} placeholder="Qty/line" onChange={(e) => updateRampUpDay(r.id, e.target.value)} className="flex-1 text-xs rounded-md border border-stone-200 px-2 py-1.5 outline-none focus:border-amber-500 tabular-nums" />
                    {noOfLines > 0 && num(r.qty) > 0 && (
                      <span className="text-[10px] text-stone-400 w-24 text-right">× {noOfLines} = {money(num(r.qty) * noOfLines)}</span>
                    )}
                    {rampPerLineTarget > 0 && num(r.qty) > 0 && (
                      <span className="text-[10px] text-stone-400 w-10 text-right">{((num(r.qty) / rampPerLineTarget) * 100).toFixed(0)}%</span>
                    )}
                    <button onClick={() => removeRampUpDay(r.id)} className="text-stone-300 active:text-red-600 shrink-0"><X size={14} /></button>
                  </div>
                ))}
                <button onClick={addRampUpDay} className="flex items-center gap-1 text-xs font-semibold text-amber-600 active:text-amber-700">
                  <Plus size={13} /> Add ramp-up day
                </button>
              </div>

              {!(num(style.orderQty) > 0) ? (
                <p className="text-xs text-stone-400">Set Order Qty in the Order Book to see the auto-calculated remaining days.</p>
              ) : !(rampTotalTarget > 0) ? (
                <p className="text-xs text-stone-400">Set Target Qty/Day/Line and No. of Lines above — remaining days after ramp-up are calculated at that combined Line Target.</p>
              ) : (
                <>
                  <div className="grid grid-cols-3 gap-2.5 mb-2.5">
                    <Stat label="Ramp-up Qty Entered" value={money(rampUpPlan.manualSum)} sub={`Day 1-${(style.obRampUpDays || []).length}, × ${noOfLines || 1} lines`} />
                    <Stat label="Remaining Qty" value={money(rampUpPlan.remaining)} sub="after ramp-up days" />
                    <Stat label="Days Needed" value={rampUpPlan.autoDaysExact.toFixed(1)} sub="@ Total Line Target/day" accent="text-amber-600" />
                  </div>
                  <div className="flex items-center justify-between rounded-lg bg-amber-50 border border-amber-200 px-3 py-2.5 mb-3">
                    <span className="text-xs font-semibold text-amber-700 uppercase tracking-wide">Lead Time — Total Days to Complete Order</span>
                    <span className="text-lg font-bold text-amber-700 tabular-nums">{rampUpPlan.totalDaysExact.toFixed(1)}</span>
                  </div>
                  <div className="space-y-1">
                    {rampUpPlan.rows.map((r) => {
                      const dayEff = rampDayEffPct(r.qty);
                      return (
                        <div key={r.day} className={`flex items-center justify-between rounded-lg border px-2.5 py-1.5 ${r.manual ? "bg-stone-50 border-stone-200" : "bg-white border-dashed border-stone-300"}`}>
                          <span className="text-xs font-semibold text-stone-600 w-16">Day {r.day}</span>
                          <span className="text-xs tabular-nums text-stone-700">{r.qty} pcs</span>
                          {rampUpTotalMp > 0 && (
                            <span
                              className={`text-[10px] font-bold tabular-nums px-1.5 py-0.5 rounded ${
                                dayEff >= 95 && dayEff <= 110 ? "bg-emerald-100 text-emerald-700" : dayEff > 110 ? "bg-red-100 text-red-700" : "bg-amber-100 text-amber-700"
                              }`}
                              title="Per-day efficiency = Qty ÷ ((Total Manpower × Working Minutes) ÷ SAM) × 100"
                            >
                              {dayEff.toFixed(0)}% eff
                            </span>
                          )}
                          <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded ${r.manual ? "bg-stone-200 text-stone-600" : "bg-amber-100 text-amber-700"}`}>
                            {r.manual ? "manual" : "auto @ Line Target"}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                  <p className="mt-1.5 text-[10px] text-stone-400">
                    Days after your manual ramp-up entries are auto-filled at the Total Line Target (Per-Line Target × No. of Lines) until the full Order Qty is covered — lead time shown as exact days (e.g. 18.5), not rounded up.
                  </p>
                </>
              )}
            </div>
          </Section>
          )}
        </>
      )}
    </div>
  );
}

// ---------------- GSD Library: Operations / Macros / Features (element/TMU engineering) ----------------
function GsdOpRow({ op, expanded, onToggle, onChange, onDelete }) {
  const r = computeGsdOp(op);
  const addElement = () => onChange({ ...op, elements: [...op.elements, newGsdElement("lib")] });
  const updateElement = (id, patch) => onChange({ ...op, elements: op.elements.map((e) => (e.id === id ? { ...e, ...patch } : e)) });
  const removeElement = (id) => onChange({ ...op, elements: op.elements.filter((e) => e.id !== id) });

  return (
    <div className="rounded-xl border border-stone-200 bg-white overflow-hidden">
      <button onClick={onToggle} className="w-full flex items-center justify-between gap-2 px-3 py-2.5 text-left">
        <div className="flex items-center gap-2 min-w-0">
          {expanded ? <ChevronDown size={14} className="text-stone-400 shrink-0" /> : <ChevronRight size={14} className="text-stone-400 shrink-0" />}
          <div className="min-w-0">
            <p className="text-xs font-bold text-stone-700 truncate">{op.code} — {op.description || op.component || "Untitled"}</p>
            <p className="text-[10px] text-stone-400">{op.mctype} · {op.elements.length} element(s)</p>
          </div>
        </div>
        <span className="text-sm font-bold text-teal-700 tabular-nums shrink-0">{r.smv.toFixed(3)}</span>
      </button>
      {expanded && (
        <div className="border-t border-stone-100 p-3 space-y-3 bg-stone-50/60">
          <div className="grid grid-cols-2 gap-2">
            <Field label="Code" value={op.code} onChange={(v) => onChange({ ...op, code: v })} type="text" />
            <Field label="Component" value={op.component} onChange={(v) => onChange({ ...op, component: v })} type="text" />
          </div>
          <Field label="Description" value={op.description} onChange={(v) => onChange({ ...op, description: v })} type="text" />
          <div className="grid grid-cols-3 gap-2">
            <label className="flex flex-col gap-1">
              <span className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">M/C Type</span>
              <select value={op.mctype} onChange={(e) => onChange({ ...op, mctype: e.target.value })} className="rounded-xl border border-stone-200 bg-white px-2 py-2 text-xs outline-none focus:border-teal-500">
                {GSD_MC_TYPES.map((m) => <option key={m} value={m}>{m}</option>)}
              </select>
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">Seam</span>
              <select value={op.seamtype} onChange={(e) => onChange({ ...op, seamtype: e.target.value })} className="rounded-xl border border-stone-200 bg-white px-2 py-2 text-xs outline-none focus:border-teal-500">
                {GSD_SEAM_TYPES.map((m) => <option key={m} value={m}>{m}</option>)}
              </select>
            </label>
            <label className="flex flex-col gap-1">
              <span className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">Skill</span>
              <select value={op.skill} onChange={(e) => onChange({ ...op, skill: e.target.value })} className="rounded-xl border border-stone-200 bg-white px-2 py-2 text-xs outline-none focus:border-teal-500">
                {GSD_SKILLS.map((m) => <option key={m} value={m}>{m}</option>)}
              </select>
            </label>
          </div>
          <div className="grid grid-cols-4 gap-2">
            <Field label="Stitch/cm" value={op.stitchespercm} onChange={(v) => onChange({ ...op, stitchespercm: v })} />
            <Field label="RPM" value={op.rpm} onChange={(v) => onChange({ ...op, rpm: v })} />
            <Field label="Bundle (min)" value={op.bundle} onChange={(v) => onChange({ ...op, bundle: v })} />
            <Field label="Allowance" value={op.allowance} onChange={(v) => onChange({ ...op, allowance: v })} suffix="%" />
          </div>

          <div>
            <div className="flex items-center justify-between mb-1.5">
              <span className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">Elements</span>
              <button onClick={addElement} className="flex items-center gap-1 text-[11px] font-bold text-teal-700"><Plus size={12} /> Add element</button>
            </div>
            <div className="space-y-1.5">
              {op.elements.map((el) => {
                const tmu = gsdElementTMU(el, op);
                return (
                  <div key={el.id} className="rounded-lg border border-stone-200 bg-white p-2 space-y-1.5">
                    <div className="flex items-center gap-1.5">
                      <select value={el.type} onChange={(e) => { const t = e.target.value; updateElement(el.id, { type: t, ...(t === "lib" ? { code: "MG2T", cm: 0 } : t === "sew" ? { tension: "N", stop: "A" } : {}) }); }} className="flex-1 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none">
                        <option value="lib">Library motion</option>
                        <option value="sew">Machine sewing</option>
                        <option value="mn">Manual entry</option>
                      </select>
                      <span className="text-[11px] font-bold text-teal-700 tabular-nums shrink-0 w-12 text-right">{tmu.toFixed(1)}</span>
                      <button onClick={() => removeElement(el.id)} className="text-stone-300 active:text-red-500 shrink-0"><X size={14} /></button>
                    </div>
                    {el.type === "lib" && (
                      <div className="flex gap-1.5">
                        <select value={el.code} onChange={(e) => { const code = e.target.value; const keys = Object.keys(GSD_CODE_LIBRARY[code].tmu).map(Number); updateElement(el.id, { code, cm: keys[0] }); }} className="flex-1 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none">
                          {Object.keys(GSD_CODE_LIBRARY).map((c) => <option key={c} value={c}>{c} — {GSD_CODE_LIBRARY[c].desc}</option>)}
                        </select>
                        <select value={el.cm} onChange={(e) => updateElement(el.id, { cm: Number(e.target.value) })} className="w-20 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none">
                          {Object.keys(GSD_CODE_LIBRARY[el.code]?.tmu || {}).map(Number).sort((a, b) => a - b).map((cm) => <option key={cm} value={cm}>{cm} cm</option>)}
                        </select>
                      </div>
                    )}
                    {el.type === "sew" && (
                      <div className="flex gap-1.5">
                        <select value={el.tension} onChange={(e) => updateElement(el.id, { tension: e.target.value })} className="flex-1 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none">
                          {Object.entries(GSD_TENSION).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                        </select>
                        <select value={el.stop} onChange={(e) => updateElement(el.id, { stop: e.target.value })} className="w-24 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none">
                          {Object.entries(GSD_STOP).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                        </select>
                        <input type="number" step="0.1" value={el.cm} onChange={(e) => updateElement(el.id, { cm: Number(e.target.value) })} placeholder="length cm" className="w-20 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none" />
                      </div>
                    )}
                    {el.type === "mn" && (
                      <input type="number" step="0.1" value={el.cm} onChange={(e) => updateElement(el.id, { cm: Number(e.target.value) })} placeholder="TMU value" className="w-24 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none" />
                    )}
                    <div className="flex items-center gap-1.5">
                      <span className="text-[10px] text-stone-400">Freq</span>
                      <input type="number" value={el.freq} onChange={(e) => updateElement(el.id, { freq: Number(e.target.value) })} className="w-14 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none" />
                      <span className="text-[10px] text-stone-400 truncate">{gsdElementDesc(el)}</span>
                    </div>
                  </div>
                );
              })}
              {op.elements.length === 0 && <p className="text-[11px] text-stone-400">No elements yet — add one above.</p>}
            </div>
          </div>

          <div className="grid grid-cols-4 gap-2 rounded-xl bg-teal-50 border border-teal-200 p-2.5">
            <Stat label="Manual min" value={r.manualMin.toFixed(3)} />
            <Stat label="Sew min" value={r.sewMin.toFixed(3)} />
            <Stat label="Basic min" value={r.basicMin.toFixed(3)} />
            <Stat label="SMV" value={r.smv.toFixed(3)} accent="text-teal-700" />
          </div>

          <button onClick={onDelete} className="flex items-center gap-1.5 text-xs font-bold text-red-500"><Trash2 size={13} /> Delete Operation</button>
        </div>
      )}
    </div>
  );
}

function GsdMacroRow({ macro, expanded, onToggle, onChange, onDelete, operations, findOp }) {
  const smv = gsdMacroSmv(macro, findOp);
  const [addOpId, setAddOpId] = useState("");
  const addItem = () => {
    if (!addOpId) return;
    onChange({ ...macro, items: [...macro.items, { id: uid(), refId: addOpId, qty: "1" }] });
    setAddOpId("");
  };
  const updateItem = (id, patch) => onChange({ ...macro, items: macro.items.map((it) => (it.id === id ? { ...it, ...patch } : it)) });
  const removeItem = (id) => onChange({ ...macro, items: macro.items.filter((it) => it.id !== id) });

  return (
    <div className="rounded-xl border border-stone-200 bg-white overflow-hidden">
      <button onClick={onToggle} className="w-full flex items-center justify-between gap-2 px-3 py-2.5 text-left">
        <div className="flex items-center gap-2 min-w-0">
          {expanded ? <ChevronDown size={14} className="text-stone-400 shrink-0" /> : <ChevronRight size={14} className="text-stone-400 shrink-0" />}
          <div className="min-w-0">
            <p className="text-xs font-bold text-stone-700 truncate">{macro.code} — {macro.name || "Untitled"}</p>
            <p className="text-[10px] text-stone-400">{macro.items.length} operation(s)</p>
          </div>
        </div>
        <span className="text-sm font-bold text-teal-700 tabular-nums shrink-0">{smv.toFixed(3)}</span>
      </button>
      {expanded && (
        <div className="border-t border-stone-100 p-3 space-y-2.5 bg-stone-50/60">
          <Field label="Macro Name" value={macro.name} onChange={(v) => onChange({ ...macro, name: v })} type="text" />
          <div className="space-y-1.5">
            {macro.items.map((it) => {
              const op = findOp(it.refId);
              return (
                <div key={it.id} className="flex items-center gap-1.5 rounded-lg border border-stone-200 bg-white px-2 py-1.5">
                  <span className="flex-1 text-[11px] text-stone-600 truncate">{op ? `${op.code} — ${op.description || op.component}` : "(missing)"}</span>
                  <input type="number" value={it.qty} onChange={(e) => updateItem(it.id, { qty: e.target.value })} className="w-12 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none tabular-nums" />
                  <button onClick={() => removeItem(it.id)} className="text-stone-300 active:text-red-500"><X size={14} /></button>
                </div>
              );
            })}
          </div>
          <div className="flex gap-1.5">
            <select value={addOpId} onChange={(e) => setAddOpId(e.target.value)} className="flex-1 rounded-xl border border-stone-200 bg-white px-2 py-2 text-xs outline-none">
              <option value="">Add operation…</option>
              {operations.map((o) => <option key={o.id} value={o.id}>{o.code} — {o.description || o.component}</option>)}
            </select>
            <button onClick={addItem} className="rounded-xl bg-teal-600 text-white px-3 py-2 text-xs font-bold"><Plus size={14} /></button>
          </div>
          <button onClick={onDelete} className="flex items-center gap-1.5 text-xs font-bold text-red-500"><Trash2 size={13} /> Delete Macro</button>
        </div>
      )}
    </div>
  );
}

function GsdFeatureRow({ feature, expanded, onToggle, onChange, onDelete, operations, macros, findOp, findMacro }) {
  const smv = gsdFeatureSmv(feature, findOp, findMacro);
  const [addType, setAddType] = useState("op");
  const [addRefId, setAddRefId] = useState("");
  const addItem = () => {
    if (!addRefId) return;
    onChange({ ...feature, items: [...feature.items, { id: uid(), refType: addType, refId: addRefId, qty: "1" }] });
    setAddRefId("");
  };
  const updateItem = (id, patch) => onChange({ ...feature, items: feature.items.map((it) => (it.id === id ? { ...it, ...patch } : it)) });
  const removeItem = (id) => onChange({ ...feature, items: feature.items.filter((it) => it.id !== id) });

  return (
    <div className="rounded-xl border border-stone-200 bg-white overflow-hidden">
      <button onClick={onToggle} className="w-full flex items-center justify-between gap-2 px-3 py-2.5 text-left">
        <div className="flex items-center gap-2 min-w-0">
          {expanded ? <ChevronDown size={14} className="text-stone-400 shrink-0" /> : <ChevronRight size={14} className="text-stone-400 shrink-0" />}
          <div className="min-w-0">
            <p className="text-xs font-bold text-stone-700 truncate">{feature.code} — {feature.name || "Untitled"}</p>
            <p className="text-[10px] text-stone-400">{feature.items.length} item(s)</p>
          </div>
        </div>
        <span className="text-sm font-bold text-teal-700 tabular-nums shrink-0">{smv.toFixed(3)}</span>
      </button>
      {expanded && (
        <div className="border-t border-stone-100 p-3 space-y-2.5 bg-stone-50/60">
          <Field label="Feature Name" value={feature.name} onChange={(v) => onChange({ ...feature, name: v })} type="text" />
          <div className="space-y-1.5">
            {feature.items.map((it) => {
              const ref = it.refType === "macro" ? findMacro(it.refId) : findOp(it.refId);
              const label = it.refType === "macro" ? (ref ? `${ref.code} — ${ref.name} (Macro)` : "(missing)") : (ref ? `${ref.code} — ${ref.description || ref.component}` : "(missing)");
              return (
                <div key={it.id} className="flex items-center gap-1.5 rounded-lg border border-stone-200 bg-white px-2 py-1.5">
                  <span className="flex-1 text-[11px] text-stone-600 truncate">{label}</span>
                  <input type="number" value={it.qty} onChange={(e) => updateItem(it.id, { qty: e.target.value })} className="w-12 rounded-md border border-stone-200 px-1.5 py-1 text-[11px] outline-none tabular-nums" />
                  <button onClick={() => removeItem(it.id)} className="text-stone-300 active:text-red-500"><X size={14} /></button>
                </div>
              );
            })}
          </div>
          <div className="flex gap-1.5">
            <select value={addType} onChange={(e) => { setAddType(e.target.value); setAddRefId(""); }} className="w-24 rounded-xl border border-stone-200 bg-white px-2 py-2 text-xs outline-none">
              <option value="op">Operation</option>
              <option value="macro">Macro</option>
            </select>
            <select value={addRefId} onChange={(e) => setAddRefId(e.target.value)} className="flex-1 rounded-xl border border-stone-200 bg-white px-2 py-2 text-xs outline-none">
              <option value="">Select…</option>
              {(addType === "op" ? operations : macros).map((o) => <option key={o.id} value={o.id}>{o.code} — {o.description || o.component || o.name}</option>)}
            </select>
            <button onClick={addItem} className="rounded-xl bg-teal-600 text-white px-3 py-2 text-xs font-bold"><Plus size={14} /></button>
          </div>
          <button onClick={onDelete} className="flex items-center gap-1.5 text-xs font-bold text-red-500"><Trash2 size={13} /> Delete Feature</button>
        </div>
      )}
    </div>
  );
}

function GsdLibraryScreen() {
  const { lib, save, findOp, findMacro } = useGsdLibrary();
  const [tab, setTab] = useState("op");
  const [search, setSearch] = useState("");
  const [expandedId, setExpandedId] = useState("");

  const setOperations = (operations) => save({ ...lib, operations });
  const setMacros = (macros) => save({ ...lib, macros });
  const setFeatures = (features) => save({ ...lib, features });

  const addOperation = () => { const op = newGsdOperation(lib.operations.length); setOperations([...lib.operations, op]); setExpandedId(op.id); setTab("op"); };
  const addMacro = () => { const m = newGsdMacro(lib.macros.length); setMacros([...lib.macros, m]); setExpandedId(m.id); setTab("macro"); };
  const addFeature = () => { const f = newGsdFeature(lib.features.length); setFeatures([...lib.features, f]); setExpandedId(f.id); setTab("feature"); };

  const q = search.toLowerCase().trim();
  const filteredOps = lib.operations.filter((o) => !q || `${o.code} ${o.description} ${o.component}`.toLowerCase().includes(q));
  const filteredMacros = lib.macros.filter((m) => !q || `${m.code} ${m.name}`.toLowerCase().includes(q));
  const filteredFeatures = lib.features.filter((f) => !q || `${f.code} ${f.name}`.toLowerCase().includes(q));

  const exportOpsCsv = () => downloadCsv("gsd_operations.csv",
    [["Code", (o) => o.code], ["Description", (o) => o.description || o.component], ["M/C", (o) => o.mctype], ["Seam (cm)", (o) => computeGsdOp(o).sewLen.toFixed(1)], ["SMV", (o) => computeGsdOp(o).smv.toFixed(3)]],
    filteredOps);
  const exportMacrosCsv = () => downloadCsv("gsd_macros.csv",
    [["Code", (m) => m.code], ["Name", (m) => m.name], ["Items", (m) => m.items.length], ["SMV", (m) => gsdMacroSmv(m, findOp).toFixed(3)]],
    filteredMacros);
  const exportFeaturesCsv = () => downloadCsv("gsd_features.csv",
    [["Code", (f) => f.code], ["Name", (f) => f.name], ["Items", (f) => f.items.length], ["SMV", (f) => gsdFeatureSmv(f, findOp, findMacro).toFixed(3)]],
    filteredFeatures);

  const gsdOpColumns = [
    ["Code", (o) => o.code], ["Component", (o) => o.component], ["Description", (o) => o.description],
    ["M/C Type", (o) => o.mctype], ["Seam", (o) => o.seamtype], ["Skill", (o) => o.skill],
    ["Stitch/cm", (o) => o.stitchespercm], ["RPM", (o) => o.rpm], ["Bundle", (o) => o.bundle], ["Allowance %", (o) => o.allowance],
    ["SMV", (o) => computeGsdOp(o).smv.toFixed(3)],
  ];
  const gsdMacroColumns = [["Code", (m) => m.code], ["Name", (m) => m.name], ["Items", (m) => m.items.length], ["SMV", (m) => gsdMacroSmv(m, findOp).toFixed(3)]];
  const gsdFeatureColumns = [["Code", (f) => f.code], ["Name", (f) => f.name], ["Items", (f) => f.items.length], ["SMV", (f) => gsdFeatureSmv(f, findOp, findMacro).toFixed(3)]];

  const handleGsdExportExcel = () => {
    if (tab === "op") downloadXlsx("gsd_operations.xlsx", "Operations", gsdOpColumns, filteredOps);
    else if (tab === "macro") downloadXlsx("gsd_macros.xlsx", "Macros", gsdMacroColumns, filteredMacros);
    else downloadXlsx("gsd_features.xlsx", "Features", gsdFeatureColumns, filteredFeatures);
  };
  const gsdFileInputRef = useRef(null);
  const [gsdSavedMsg, setGsdSavedMsg] = useState("");
  const handleGsdSave = () => {
    save({ ...lib });
    setGsdSavedMsg("Saved ✓");
    setTimeout(() => setGsdSavedMsg(""), 2000);
  };
  const handleGsdImportExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (tab !== "op") { alert("Import currently supports Operations only — add Macro/Feature items manually after import."); e.target.value = ""; return; }
    readXlsxFile(file, (rows) => {
      const imported = rows.map((r, i) => ({
        ...newGsdOperation(lib.operations.length + i),
        code: String(r["Code"] || `OP-IMP-${i + 1}`),
        component: String(r["Component"] || ""),
        description: String(r["Description"] || ""),
        mctype: String(r["M/C Type"] || GSD_MC_TYPES[14]),
        seamtype: String(r["Seam"] || GSD_SEAM_TYPES[0]),
        skill: String(r["Skill"] || GSD_SKILLS[0]),
        stitchespercm: Number(r["Stitch/cm"] || 4.5),
        rpm: Number(r["RPM"] || 4000),
        bundle: Number(r["Bundle"] || 0),
        allowance: Number(r["Allowance %"] || 15),
      }));
      setOperations([...lib.operations, ...imported]);
      setGsdSavedMsg(`Imported ${imported.length} operation(s) from Excel ✓`);
      setTimeout(() => setGsdSavedMsg(""), 3000);
    });
    e.target.value = "";
  };

  return (
    <div className="space-y-4 pb-16">
      <div className="rounded-2xl border border-stone-200 bg-white p-3 space-y-2.5">
        <div className="flex gap-1.5">
          {[["op", "Operations", lib.operations.length], ["macro", "Macros", lib.macros.length], ["feature", "Features", lib.features.length]].map(([k, l, c]) => (
            <button key={k} onClick={() => { setTab(k); setExpandedId(""); }} style={tab === k ? { backgroundColor: "#0b3d33" } : undefined} className={`flex-1 rounded-lg px-2 py-2 text-xs font-bold transition-colors ${tab === k ? "text-white" : "bg-stone-100 text-stone-500"}`}>
              {l} <span className="opacity-70">({c})</span>
            </button>
          ))}
        </div>
        <div className="flex items-center gap-1.5 rounded-xl border border-stone-200 px-2.5 py-1.5">
          <Search size={14} className="text-stone-300 shrink-0" />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search…" className="flex-1 text-xs outline-none" />
        </div>
        <div className="flex flex-wrap gap-1.5">
          <button
            onClick={tab === "op" ? addOperation : tab === "macro" ? addMacro : addFeature}
            className="flex-1 flex items-center justify-center gap-1.5 rounded-xl bg-teal-600 text-white px-3 py-2 text-xs font-bold"
          >
            <Plus size={14} /> New {tab === "op" ? "Operation" : tab === "macro" ? "Macro" : "Feature"}
          </button>
          <button onClick={handleGsdSave} className="flex items-center gap-1.5 rounded-xl bg-emerald-600 text-white px-3 py-2 text-xs font-bold">
            <Save size={14} /> Save
          </button>
          <button onClick={tab === "op" ? exportOpsCsv : tab === "macro" ? exportMacrosCsv : exportFeaturesCsv} className="flex items-center gap-1.5 rounded-xl border border-stone-200 px-3 py-2 text-xs font-bold text-stone-600">
            <Download size={14} /> CSV
          </button>
          <button onClick={handleGsdExportExcel} className="flex items-center gap-1.5 rounded-xl border border-stone-200 px-3 py-2 text-xs font-bold text-stone-600">
            <FileSpreadsheet size={14} /> Excel
          </button>
          <button onClick={() => gsdFileInputRef.current?.click()} className="flex items-center gap-1.5 rounded-xl border border-stone-200 px-3 py-2 text-xs font-bold text-stone-600">
            <Upload size={14} /> Import
          </button>
          <input ref={gsdFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleGsdImportExcel} className="hidden" />
        </div>
        {gsdSavedMsg && <p className="text-center text-xs font-semibold text-emerald-600 flex items-center justify-center gap-1"><Check size={13} /> {gsdSavedMsg}</p>}
      </div>

      <div className="space-y-1.5">
        {tab === "op" && (filteredOps.length === 0 ? (
          <EmptyState icon={Layers} title="No operations yet" body="Build reusable GSD operations (element/TMU engineered) here — use them in Macros, Features, or pull straight into any style's Operation Bulletin." actionLabel="New Operation" onAction={addOperation} />
        ) : filteredOps.map((op) => (
          <GsdOpRow key={op.id} op={op} expanded={expandedId === op.id} onToggle={() => setExpandedId(expandedId === op.id ? "" : op.id)} onChange={(next) => setOperations(lib.operations.map((o) => (o.id === op.id ? next : o)))} onDelete={() => { setOperations(lib.operations.filter((o) => o.id !== op.id)); setExpandedId(""); }} />
        )))}
        {tab === "macro" && (filteredMacros.length === 0 ? (
          <EmptyState icon={Layers} title="No macros yet" body="A Macro groups several Operations that always happen together (e.g. a full sleeve set)." actionLabel="New Macro" onAction={addMacro} />
        ) : filteredMacros.map((m) => (
          <GsdMacroRow key={m.id} macro={m} expanded={expandedId === m.id} onToggle={() => setExpandedId(expandedId === m.id ? "" : m.id)} onChange={(next) => setMacros(lib.macros.map((x) => (x.id === m.id ? next : x)))} onDelete={() => { setMacros(lib.macros.filter((x) => x.id !== m.id)); setExpandedId(""); }} operations={lib.operations} findOp={findOp} />
        )))}
        {tab === "feature" && (filteredFeatures.length === 0 ? (
          <EmptyState icon={Layers} title="No features yet" body="A Feature groups Macros/Operations into a full garment feature (e.g. Placket, Collar) for fast style build-up." actionLabel="New Feature" onAction={addFeature} />
        ) : filteredFeatures.map((f) => (
          <GsdFeatureRow key={f.id} feature={f} expanded={expandedId === f.id} onToggle={() => setExpandedId(expandedId === f.id ? "" : f.id)} onChange={(next) => setFeatures(lib.features.map((x) => (x.id === f.id ? next : x)))} onDelete={() => { setFeatures(lib.features.filter((x) => x.id !== f.id)); setExpandedId(""); }} operations={lib.operations} macros={lib.macros} findOp={findOp} findMacro={findMacro} />
        )))}
      </div>
    </div>
  );
}

function greetingWord() {
  const h = new Date().getHours();
  if (h < 12) return "Good morning";
  if (h < 17) return "Good afternoon";
  return "Good evening";
}
function DashboardScreen({ styles, loginUser, onNavigate }) {
  const totalOps = styles.reduce((a, s) => a + (s.obOperations?.length || 0), 0);
  const totalSmv = styles.reduce((a, s) => a + (s.obOperations || []).reduce((x, o) => x + num(o.sam), 0), 0);
  const buyers = new Set(styles.map((s) => s.buyer).filter(Boolean)).size;
  const cards = [
    { label: "Active Styles", value: styles.length, icon: Boxes, bg: "bg-teal-500" },
    { label: "Buyers", value: buyers, icon: Users, bg: "bg-emerald-500" },
    { label: "Total Operations", value: totalOps, icon: ClipboardList, bg: "bg-violet-500" },
    { label: "Total SMV", value: totalSmv.toFixed(2), icon: Gauge, bg: "bg-amber-500" },
    { label: "Total Users", value: 1, icon: Users, bg: "bg-stone-500" },
  ];
  return (
    <div className="space-y-4 pb-16">
      <div className="rounded-2xl p-6" style={{ background: "linear-gradient(160deg, #0b3d33 0%, #0f4d40 55%, #116857 100%)" }}>
        <h2 className="text-xl font-extrabold text-white">{greetingWord()}, {loginUser || "Admin"}!</h2>
        <p className="text-sm text-emerald-100/80 mt-1.5 max-w-sm">Here's what's happening with your Operation Bulletin data today.</p>
        <button onClick={() => onNavigate("ob")} className="mt-4 flex items-center gap-1.5 rounded-xl bg-white/15 hover:bg-white/25 text-white font-bold text-sm px-4 py-2.5 transition-colors">
          Open Operation Bulletin <ChevronRight size={16} />
        </button>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
        {cards.map((c) => (
          <div key={c.label} className="rounded-2xl bg-white border border-stone-200 p-4 flex items-center gap-3">
            <span className={`h-11 w-11 rounded-xl ${c.bg} flex items-center justify-center shrink-0`}>
              <c.icon size={20} className="text-white" />
            </span>
            <div className="min-w-0">
              <p className="text-[11px] text-stone-400 truncate">{c.label}</p>
              <p className="text-xl font-extrabold text-stone-800 tabular-nums">{c.value}</p>
            </div>
          </div>
        ))}
      </div>

      <div className="rounded-2xl bg-white border border-stone-200 p-4">
        <p className="text-[11px] font-extrabold tracking-wider text-stone-400 uppercase mb-2.5">Quick Actions</p>
        <div className="grid grid-cols-2 gap-2.5">
          <button onClick={() => onNavigate("ob")} className="flex items-center gap-2 rounded-xl border border-stone-200 px-3 py-2.5 text-left hover:bg-stone-50">
            <ClipboardList size={16} className="text-amber-600 shrink-0" />
            <span className="text-xs font-bold text-stone-700">Operation Bulletin</span>
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------- Standalone App shell: loads styles from the shared SGPS storage ----------------
// Admin screen: create new users with granular menu/factory access, and manage existing ones.
// Only rendered/reachable when the logged-in user is an Admin (see sidebar + mainView gating in App).
function AdminScreen({ users, onSaveUsers, factories, currentUser, menuOptions }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [isAdminNew, setIsAdminNew] = useState(false);
  const [menus, setMenus] = useState([]);
  const [factoryIds, setFactoryIds] = useState([]);
  const [error, setError] = useState("");

  const toggleMenu = (key) => {
    setMenus((prev) => (prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]));
  };
  const toggleFactory = (id) => {
    setFactoryIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  };

  const handleCreate = () => {
    const uname = username.trim();
    if (!uname) { setError("Enter a username."); return; }
    if (!password) { setError("Set a password."); return; }
    if (users.some((u) => u.username.toLowerCase() === uname.toLowerCase())) {
      setError("That username is already taken.");
      return;
    }
    const newUser = {
      id: uid(),
      username: uname,
      password,
      isAdmin: isAdminNew,
      menus: isAdminNew ? menuOptions.map((m) => m.key) : menus,
      factoryIds: isAdminNew ? factories.map((f) => f.id) : factoryIds,
    };
    onSaveUsers([...users, newUser]);
    setUsername(""); setPassword(""); setIsAdminNew(false); setMenus([]); setFactoryIds([]); setError("");
  };

  const adminCount = users.filter((u) => u.isAdmin).length;

  const handleDelete = (id) => {
    if (id === currentUser?.id) return; // can't delete yourself while logged in
    const target = users.find((u) => u.id === id);
    if (target?.isAdmin && adminCount <= 1) return; // always keep at least one admin
    onSaveUsers(users.filter((u) => u.id !== id));
  };

  const handlePasswordChange = (id, next) => {
    onSaveUsers(users.map((u) => (u.id === id ? { ...u, password: next } : u)));
  };

  return (
    <div className="space-y-4">
      <div className="bg-white rounded-2xl border border-stone-200 shadow-sm p-4">
        <h2 className="flex items-center gap-1.5 text-sm font-extrabold text-stone-800 mb-4">
          <Plus size={16} /> Create New User
        </h2>
        {error && <div className="mb-3 text-[12px] font-semibold text-red-600">{error}</div>}
        <div className="grid sm:grid-cols-2 gap-3 mb-3">
          <div>
            <label className="block text-[10.5px] uppercase tracking-wide text-stone-500 font-bold mb-1.5">Username</label>
            <input
              value={username} onChange={(e) => setUsername(e.target.value)} placeholder="e.g. cutting_incharge"
              className="w-full px-3 py-2.5 rounded-lg border border-stone-200 bg-stone-50 text-sm outline-none focus:border-[#116857] focus:bg-white"
            />
          </div>
          <div>
            <label className="block text-[10.5px] uppercase tracking-wide text-stone-500 font-bold mb-1.5">Password</label>
            <input
              value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Set a password"
              className="w-full px-3 py-2.5 rounded-lg border border-stone-200 bg-stone-50 text-sm outline-none focus:border-[#116857] focus:bg-white"
            />
          </div>
        </div>

        <label className="flex items-center gap-2 mb-4 cursor-pointer">
          <input type="checkbox" checked={isAdminNew} onChange={(e) => setIsAdminNew(e.target.checked)} className="w-4 h-4 accent-amber-500" />
          <span className="text-xs font-bold text-stone-700">Admin (full access to every menu)</span>
        </label>

        {!isAdminNew && factories.length > 0 && (
          <div className="mb-4">
            <p className="text-[10.5px] uppercase tracking-wide text-stone-500 font-bold mb-2">Factory Access</p>
            <div className="flex flex-wrap gap-x-5 gap-y-1.5">
              {factories.map((f) => (
                <label key={f.id} className="flex items-center gap-1.5 cursor-pointer">
                  <input type="checkbox" checked={factoryIds.includes(f.id)} onChange={() => toggleFactory(f.id)} className="w-4 h-4 accent-amber-500" />
                  <span className="text-xs font-semibold text-stone-700">{f.name || f.code || "Factory"}</span>
                </label>
              ))}
            </div>
          </div>
        )}

        {!isAdminNew && (
          <div className="mb-4">
            <p className="text-[10.5px] uppercase tracking-wide text-stone-500 font-bold mb-2">Menu Access</p>
            <div className="grid grid-cols-2 gap-x-5 gap-y-1.5">
              {menuOptions.map((m) => (
                <label key={m.key} className="flex items-center gap-1.5 cursor-pointer">
                  <input type="checkbox" checked={menus.includes(m.key)} onChange={() => toggleMenu(m.key)} className="w-4 h-4 accent-amber-500" />
                  <span className="text-xs font-semibold text-stone-700">{m.label}</span>
                </label>
              ))}
            </div>
          </div>
        )}

        <button
          onClick={handleCreate}
          className="w-full py-2.5 rounded-lg font-extrabold text-sm bg-amber-400 text-stone-900 hover:brightness-105 flex items-center justify-center gap-1.5"
        >
          <Plus size={15} /> Create User
        </button>
      </div>

      <div className="bg-white rounded-2xl border border-stone-200 shadow-sm p-4">
        <h2 className="flex items-center gap-1.5 text-sm font-extrabold text-stone-800 mb-3">
          <Users size={16} /> Users ({users.length})
        </h2>
        <div className="space-y-2.5">
          {users.map((u) => {
            const canDelete = u.id !== currentUser?.id && !(u.isAdmin && adminCount <= 1);
            return (
              <div key={u.id} className="rounded-xl border border-stone-200 p-3">
                <div className="flex items-center gap-2 flex-wrap mb-1.5">
                  <span className="text-sm font-extrabold text-stone-800">{u.username}</span>
                  {u.isAdmin && (
                    <span className="text-[10px] font-extrabold px-2 py-0.5 rounded-full bg-amber-100 text-amber-700 border border-amber-300">ADMIN</span>
                  )}
                  {u.id === currentUser?.id && <span className="text-[11px] text-stone-400">(you)</span>}
                  {canDelete && (
                    <button onClick={() => handleDelete(u.id)} className="ml-auto text-stone-400 hover:text-red-600" aria-label="Delete user">
                      <Trash2 size={15} />
                    </button>
                  )}
                </div>
                <div className="flex items-center gap-2 mb-1.5">
                  <span className="text-[11px] text-stone-500 font-semibold">Password:</span>
                  <input
                    defaultValue={u.password}
                    onBlur={(e) => handlePasswordChange(u.id, e.target.value)}
                    className="px-2.5 py-1.5 rounded-md border border-stone-200 bg-stone-50 text-xs outline-none focus:border-[#116857] focus:bg-white"
                  />
                </div>
                <p className="text-[11px] text-stone-400">
                  {u.isAdmin
                    ? "Full access to every menu and factory"
                    : `Menus: ${(u.menus || []).map((k) => menuOptions.find((m) => m.key === k)?.label || k).join(", ") || "None"}`}
                </p>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function App() {
  const [factories, setFactories] = useState([]);
  const [currentFactoryId, setCurrentFactoryId] = useState(null);
  const [allStyles, setAllStyles] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [menuSearch, setMenuSearch] = useState("");
  const { users, save: saveUsers } = useObUsers();

  const [authed, setAuthed] = useState(() => {
    try { return sessionStorage.getItem(OB_SESSION_KEY) === "1"; } catch (e) { return false; }
  });
  const [currentUser, setCurrentUser] = useState(() => {
    try {
      if (sessionStorage.getItem(OB_SESSION_KEY) !== "1") return null;
      const uname = sessionStorage.getItem(OB_SESSION_USER_KEY);
      return loadObUsersRaw().find((u) => u.username === uname) || null;
    } catch (e) { return null; }
  });
  const [loginUser, setLoginUser] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [loginError, setLoginError] = useState("");
  const [mainView, setMainView] = useState("dashboard"); // 'dashboard' | 'ob' | 'admin'
  const [obTab, setObTab] = useState("target"); // sub-tab within Operation Bulletin
  const [obMenuOpen, setObMenuOpen] = useState(true); // accordion state for OB submenu

  const isAdmin = !!currentUser?.isAdmin;
  const allowedMenus = isAdmin ? OB_MENU_OPTIONS.map((m) => m.key) : (currentUser?.menus || []);
  const canSeeMenu = (key) => isAdmin || allowedMenus.includes(key);
  // "Sub Entries", "OB Template Upload" and "Saved OB (Buyer-wise)" now live as their own
  // top-level main-menu buttons (see sidebar below) instead of being nested inside the
  // Operation Bulletin accordion.
  const OB_SUB_TABS = [
    ["newstyle", "+ New Style"],
    ["target", "Target & Manpower"],
    ["ops", "Operations"],
    ["balance", "SMV & Balance"],
    ["rampup", "Ramp-up"],
  ];
  const obVisibleTabs = OB_SUB_TABS.filter(([k]) => canSeeMenu(k));

  const handleLogin = (e) => {
    e.preventDefault();
    const match = users.find((u) => u.username === loginUser.trim() && u.password === loginPass);
    if (match) {
      try {
        sessionStorage.setItem(OB_SESSION_KEY, "1");
        sessionStorage.setItem(OB_SESSION_USER_KEY, match.username);
      } catch (err) {}
      setLoginError("");
      setCurrentUser(match);
      setAuthed(true);
    } else {
      setLoginError("Incorrect username or password.");
    }
  };

  const handleLogout = () => {
    try {
      sessionStorage.removeItem(OB_SESSION_KEY);
      sessionStorage.removeItem(OB_SESSION_USER_KEY);
    } catch (err) {}
    setLoginUser("");
    setLoginPass("");
    setCurrentUser(null);
    setAuthed(false);
    setMainView("dashboard");
  };

  // If a session cookie says "authed" but the user was since deleted from Admin, force re-login.
  useEffect(() => {
    if (authed && !currentUser) handleLogout();
  }, [authed, currentUser]);

  // Right after login (or on a restored session), make sure the user lands on a screen they
  // actually have access to instead of defaulting to a Dashboard they may not be allowed to see.
  useEffect(() => {
    if (!authed || !currentUser || isAdmin) return;
    if (canSeeMenu(mainView) || (mainView === "ob" && (obVisibleTabs.length > 0 || canSeeMenu("subentries") || canSeeMenu("templateupload") || canSeeMenu("savedob")))) return;
    if (canSeeMenu("dashboard")) { setMainView("dashboard"); return; }
    if (obVisibleTabs.length > 0) { setMainView("ob"); setObTab(obVisibleTabs[0][0]); return; }
    if (canSeeMenu("subentries")) { setMainView("ob"); setObTab("subentries"); return; }
    if (canSeeMenu("templateupload")) { setMainView("ob"); setObTab("templateupload"); return; }
    if (canSeeMenu("savedob")) { setMainView("ob"); setObTab("savedob"); return; }
  }, [authed, currentUser]);

  useEffect(() => {
    try {
      const facsRaw = localStorage.getItem("sgps-factories");
      const facs = facsRaw ? JSON.parse(facsRaw) : [];
      setFactories(facs);
      const stylesRaw = localStorage.getItem("sgps-styles");
      setAllStyles(stylesRaw ? JSON.parse(stylesRaw) : []);
      const facIdRaw = localStorage.getItem("sgps-current-factory");
      const facId = facIdRaw ? JSON.parse(facIdRaw) : null;
      setCurrentFactoryId(facId && facs.some((f) => f.id === facId) ? facId : (facs[0]?.id || null));
    } catch (e) {
      // ignore — starts empty if SGPS hasn't been opened on this device/browser yet
    }
    setLoaded(true);
  }, []);

  const styles = useMemo(
    () => (currentFactoryId ? allStyles.filter((s) => s.factoryId === currentFactoryId) : allStyles),
    [allStyles, currentFactoryId]
  );
  const currentFactory = factories.find((f) => f.id === currentFactoryId) || factories[0];

  // Patches one style in place and writes the FULL styles array back to the same "sgps-styles"
  // key the main SGPS app reads — so nothing else in SGPS is touched, only this style's fields.
  const updateStyle = (updated) => {
    setAllStyles((prev) => {
      const next = prev.map((s) => (s.id === updated.id ? updated : s));
      try {
        localStorage.setItem("sgps-styles", JSON.stringify(next));
      } catch (e) {
        // storage write failed — state still updates in-memory for this session
      }
      return next;
    });
  };

  // Creates a brand-new style directly from Operation Bulletin (styleNo/buyer/orderQty typed in),
  // writing it into the SAME "sgps-styles" array SGPS reads — so it shows up back there too.
  const createStyle = (partial) => {
    const newStyle = {
      id: uid(), styleNo: "", buyer: "", product: "", colour: "", orderQty: "",
      factoryId: currentFactoryId || undefined, obOperations: [], departments: blankDepartments(),
      ...partial,
    };
    setAllStyles((prev) => {
      const next = [...prev, newStyle];
      try { localStorage.setItem("sgps-styles", JSON.stringify(next)); } catch (e) {}
      return next;
    });
    return newStyle.id;
  };

  if (!authed) {
    return (
      <div
        className="min-h-screen flex items-center justify-center p-4"
        style={{ background: "linear-gradient(160deg, #0b3d33 0%, #0f4d40 55%, #116857 100%)" }}
      >
        <div className="w-full max-w-sm bg-white rounded-2xl shadow-2xl px-8 py-9">
          <div className="flex flex-col items-center mb-5">
            <img src={GSD_LOGO_DATA_URI} alt="Smart Data OB" className="h-16 w-16 rounded-2xl object-cover shadow-md mb-2.5" />
            <h1 className="text-base font-extrabold tracking-wide text-center" style={{ color: "#0b3d33" }}>Smart Data OB</h1>
            <p className="text-[11px] text-stone-500 text-center mt-1">Operation Bulletin · Sign in to continue</p>
          </div>
          <div onKeyDown={(e) => { if (e.key === "Enter") handleLogin(e); }}>
            <div className="mb-3.5">
              <label className="block text-[10.5px] uppercase tracking-wide text-stone-500 font-bold mb-1.5">Username</label>
              <input
                type="text" value={loginUser} onChange={(e) => setLoginUser(e.target.value)}
                autoComplete="username" placeholder="Enter username"
                className="w-full px-3 py-2.5 rounded-lg border border-stone-200 bg-stone-50 text-sm outline-none focus:border-[#116857] focus:bg-white"
              />
            </div>
            <div className="mb-1">
              <label className="block text-[10.5px] uppercase tracking-wide text-stone-500 font-bold mb-1.5">Password</label>
              <div className="relative">
                <input
                  type={showPass ? "text" : "password"} value={loginPass} onChange={(e) => setLoginPass(e.target.value)}
                  autoComplete="current-password" placeholder="Enter password"
                  className="w-full px-3 py-2.5 pr-10 rounded-lg border border-stone-200 bg-stone-50 text-sm outline-none focus:border-[#116857] focus:bg-white"
                />
                <button
                  type="button" onClick={() => setShowPass((s) => !s)}
                  aria-label={showPass ? "Hide password" : "Show password"}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-stone-400 hover:text-[#116857]"
                >
                  {showPass ? <EyeOff size={17} /> : <Eye size={17} />}
                </button>
              </div>
            </div>
            {loginError && <div className="text-[12px] text-red-600 font-semibold mt-2.5 text-center">{loginError}</div>}
            <button
              type="button" onClick={handleLogin}
              className="w-full mt-5 py-2.5 rounded-lg font-extrabold text-sm bg-amber-400 text-stone-900 hover:brightness-105"
            >
              Sign In
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (!loaded) return null;

  return (
    <div className="min-h-screen bg-stone-100">
      <header className="bg-white border-b border-stone-200 px-4 py-3 flex items-center gap-2.5">
        <button onClick={() => setMenuOpen((o) => !o)} className="h-9 w-9 rounded-lg flex items-center justify-center shrink-0 text-stone-600 hover:bg-stone-100 active:bg-stone-200">
          <MenuIcon size={20} />
        </button>
        <h1 className="text-sm font-extrabold text-stone-800 flex-1">
          {mainView === "dashboard" ? "Dashboard" : mainView === "ob" ? "Operation Bulletin" : mainView === "admin" ? "Admin" : ""}
        </h1>
        <span className="h-8 w-8 rounded-full bg-teal-100 text-xs font-extrabold flex items-center justify-center shrink-0" style={{ color: "#0b3d33" }}>
          {(currentUser?.username || "AD").slice(0, 2).toUpperCase()}
        </span>
        <ChevronDown size={14} className="text-stone-400 shrink-0" />
      </header>

      {mainView === "dashboard" && (
        <div className="bg-white border-b border-stone-200 px-4 py-2.5 flex items-center gap-6">
          <div>
            <p className="text-[10px] font-bold text-stone-400 tracking-wider">ACTIVE STYLES</p>
            <p className="text-lg font-extrabold text-stone-800 tabular-nums">{styles.length}</p>
          </div>
          <div>
            <p className="text-[10px] font-bold text-stone-400 tracking-wider">TOTAL OPERATIONS</p>
            <p className="text-lg font-extrabold text-stone-800 tabular-nums">{styles.reduce((a, s) => a + (s.obOperations?.length || 0), 0)}</p>
          </div>
        </div>
      )}

      <div className="flex" style={{ minHeight: "calc(100vh - 57px)" }}>
        {menuOpen && (
          <div className="fixed inset-0 bg-black/50 z-40 sm:hidden" onClick={() => setMenuOpen(false)} />
        )}
        <aside
          style={{ backgroundColor: "#0b3d33" }}
          className={`fixed sm:static inset-y-0 left-0 z-50 w-64 sm:w-56 shrink-0 flex flex-col py-3 px-2.5 overflow-y-auto shadow-2xl sm:shadow-none transition-transform duration-200 ${
            menuOpen ? "translate-x-0" : "-translate-x-full sm:translate-x-0"
          }`}
        >
          {/* Logo card */}
          <div className="flex items-center gap-2.5 rounded-xl px-3 py-2.5 mb-1.5 bg-white shadow-sm">
            <img src={GSD_LOGO_DATA_URI} alt="Smart Data OB" className="h-8 w-8 rounded-md object-cover shrink-0" />
            <span className="text-xs font-extrabold leading-tight" style={{ color: "#0b3d33" }}>Smart Data OB</span>
          </div>
          <p className="text-[10px] text-emerald-300/80 text-center mb-3 tracking-wide">Accurate · Right Time · Right Output</p>

          {/* Current factory pill */}
          {factories.length > 0 && (
            <div className="relative mb-3">
              <div className="flex items-center gap-2 rounded-xl bg-white/10 px-2.5 py-2">
                <span className="w-6 h-6 rounded-md bg-amber-400 text-stone-900 text-[10px] font-extrabold flex items-center justify-center shrink-0">
                  {(currentFactory?.name || currentFactory?.code || "FA").slice(0, 2).toUpperCase()}
                </span>
                <span className="flex-1 text-xs font-bold text-white truncate">{currentFactory?.name || currentFactory?.code || "Factory"}</span>
                {factories.length > 1 && <ChevronDown size={14} className="text-white/60 shrink-0" />}
              </div>
              {factories.length > 1 && (
                <select
                  value={currentFactoryId || ""}
                  onChange={(e) => {
                    setCurrentFactoryId(e.target.value);
                    try { localStorage.setItem("sgps-current-factory", JSON.stringify(e.target.value)); } catch (err) {}
                  }}
                  className="absolute inset-0 opacity-0 cursor-pointer"
                >
                  {factories.map((f) => <option key={f.id} value={f.id}>{f.name || f.code || "Factory"}</option>)}
                </select>
              )}
            </div>
          )}

          {/* Search */}
          <div className="flex items-center gap-1.5 rounded-xl bg-white/10 px-2.5 py-2 mb-3">
            <Search size={14} className="text-emerald-300/70 shrink-0" />
            <input
              value={menuSearch}
              onChange={(e) => setMenuSearch(e.target.value)}
              placeholder="Search menus…"
              className="flex-1 bg-transparent text-xs text-white placeholder-emerald-300/50 outline-none"
            />
          </div>

          {/* Menu */}
          {canSeeMenu("dashboard") && (!menuSearch || "overview".includes(menuSearch.toLowerCase()) || "dashboard".includes(menuSearch.toLowerCase())) && (
            <>
              <p className="text-[10px] font-extrabold tracking-wider text-emerald-300/70 uppercase mb-1.5 px-1">Overview</p>
              <button
                onClick={() => { setMainView("dashboard"); setMenuOpen(false); }}
                className={`w-full flex items-center gap-2.5 rounded-xl px-3 py-2.5 mb-1 text-left transition-colors ${
                  mainView === "dashboard" ? "bg-amber-900/40 text-amber-300 font-bold" : "text-emerald-100/90 hover:bg-white/10 hover:text-white"
                }`}
              >
                <Home size={17} className="shrink-0" />
                <span className="text-xs font-bold leading-tight">Dashboard</span>
              </button>
            </>
          )}

          <p className="text-[10px] font-extrabold tracking-wider text-emerald-300/70 uppercase mb-1.5 px-1 mt-1">Menu</p>

          {obVisibleTabs.length > 0 && (!menuSearch || "operation bulletin target manpower operations smv balance ramp-up".includes(menuSearch.toLowerCase())) && (
          <button
            onClick={() => {
              setMainView("ob");
              if (!obVisibleTabs.some(([k]) => k === obTab)) setObTab(obVisibleTabs[0]?.[0] || "target");
              setObMenuOpen((o) => !o);
            }}
            className="w-full flex items-center justify-between gap-1 rounded-xl px-3 py-2.5 mb-0.5 text-emerald-100/90 hover:bg-white/10 hover:text-white transition-colors"
          >
            <span className="flex items-center gap-2.5">
              <ClipboardList size={17} className="shrink-0" />
              <span className="text-xs font-bold leading-tight">Operation Bulletin</span>
            </span>
            <ChevronDown size={14} className={`shrink-0 transition-transform ${obMenuOpen ? "" : "-rotate-90"}`} />
          </button>
          )}

          {obVisibleTabs.length > 0 && obMenuOpen && (
            <div className="mb-1 flex flex-col gap-0.5 pl-2.5 py-1.5 ml-1 rounded-xl bg-black/20 border border-white/5">
              {obVisibleTabs.map(([k, l]) => (
                <button
                  key={k}
                  onClick={() => { setMainView("ob"); setObTab(k); setMenuOpen(false); }}
                  className={`rounded-lg px-2.5 py-2 text-left transition-colors ${
                    mainView === "ob" && obTab === k
                      ? "bg-amber-900/40 text-amber-300 font-bold"
                      : "text-emerald-50/70 hover:bg-white/10 hover:text-white font-semibold"
                  }`}
                >
                  <span className="text-[11px] leading-tight">{l}</span>
                </button>
              ))}
            </div>
          )}

          {canSeeMenu("subentries") && (!menuSearch || "sub entries".includes(menuSearch.toLowerCase())) && (
          <button
            onClick={() => { setMainView("ob"); setObTab("subentries"); setMenuOpen(false); }}
            className={`flex items-center gap-2.5 rounded-xl px-3 py-2.5 mb-1 text-left transition-colors ${
              mainView === "ob" && obTab === "subentries" ? "bg-amber-900/40 text-amber-300 font-bold" : "text-emerald-50/80 hover:bg-white/10 hover:text-white font-semibold"
            }`}
          >
            <Layers size={17} className="shrink-0" />
            <span className="text-xs leading-tight">Sub Entries</span>
          </button>
          )}

          {canSeeMenu("templateupload") && (!menuSearch || "ob template upload".includes(menuSearch.toLowerCase())) && (
          <button
            onClick={() => { setMainView("ob"); setObTab("templateupload"); setMenuOpen(false); }}
            className={`flex items-center gap-2.5 rounded-xl px-3 py-2.5 mb-1 text-left transition-colors ${
              mainView === "ob" && obTab === "templateupload" ? "bg-amber-900/40 text-amber-300 font-bold" : "text-emerald-50/80 hover:bg-white/10 hover:text-white font-semibold"
            }`}
          >
            <Upload size={17} className="shrink-0" />
            <span className="text-xs leading-tight">OB Template Upload</span>
          </button>
          )}

          {canSeeMenu("savedob") && (!menuSearch || "saved ob buyer-wise print share".includes(menuSearch.toLowerCase())) && (
          <button
            onClick={() => { setMainView("ob"); setObTab("savedob"); setMenuOpen(false); }}
            className={`flex items-center gap-2.5 rounded-xl px-3 py-2.5 mb-1 text-left transition-colors ${
              mainView === "ob" && obTab === "savedob" ? "bg-amber-900/40 text-amber-300 font-bold" : "text-emerald-50/80 hover:bg-white/10 hover:text-white font-semibold"
            }`}
          >
            <FolderOpen size={17} className="shrink-0" />
            <span className="text-xs leading-tight">Saved OB (Buyer-wise)</span>
          </button>
          )}

          {isAdmin && (!menuSearch || "admin users create user access".includes(menuSearch.toLowerCase())) && (
          <button
            onClick={() => { setMainView("admin"); setMenuOpen(false); }}
            className={`flex items-center gap-2.5 rounded-xl px-3 py-2.5 mb-1 text-left transition-colors ${
              mainView === "admin" ? "bg-amber-900/40 text-amber-300 font-bold" : "text-emerald-50/80 hover:bg-white/10 hover:text-white font-semibold"
            }`}
          >
            <Settings size={17} className="shrink-0" />
            <span className="text-xs leading-tight">Admin</span>
          </button>
          )}

          {/* Footer: user + logout */}
          <div className="mt-auto pt-3 border-t border-white/10 flex items-center justify-between px-1">
            <div>
              <p className="text-xs font-bold text-white leading-tight">{currentUser?.username || "admin"}</p>
              <p className="text-[10px] text-emerald-300/70 leading-tight">{isAdmin ? "Admin" : "User"}</p>
            </div>
            <button onClick={handleLogout} className="flex items-center gap-1 text-amber-300 text-xs font-bold hover:text-amber-200">
              <LogOut size={13} />
              <span>Logout</span>
            </button>
          </div>
        </aside>

        <main className={`flex-1 min-w-0 mx-auto p-4 ${mainView === "ob" ? "max-w-full" : "max-w-2xl lg:max-w-3xl xl:max-w-4xl"}`}>
          {mainView === "dashboard" ? (
            <DashboardScreen styles={styles} loginUser={currentUser?.username} onNavigate={setMainView} />
          ) : mainView === "ob" ? (
            <OperationBulletinScreen styles={styles} onUpdateStyle={updateStyle} onCreateStyle={createStyle} obTab={obTab} setObTab={setObTab} />
          ) : mainView === "admin" ? (
            isAdmin ? (
              <AdminScreen users={users} onSaveUsers={saveUsers} factories={factories} currentUser={currentUser} menuOptions={OB_MENU_OPTIONS} />
            ) : null
          ) : null}
        </main>
      </div>
    </div>

  );
}

export default App;
