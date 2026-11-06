// Week 43 (ADO, Oct 26-30 2026) styled workbook from week43final.json
const ExcelJS = require("exceljs");
const fs = require("fs");
const STORIES = JSON.parse(fs.readFileSync("week43final.json"));
const F2E = JSON.parse(fs.readFileSync("data.json"))["Features to Epic"];
const OUT = process.argv[2] || ".";

const ADO = id => `https://dev.azure.com/oneCELA/TSI%20Sales%20and%20Ops/_workitems/edit/${id}`;

// ---------- classifications ----------
const CATEGORY = {
  568362:"Planned Delivery",569088:"Planned Delivery",585785:"Run the Business",586680:"Planned Delivery",
  588673:"Planned Delivery",589085:"Planned Delivery",589105:"Run the Business",589106:"Run the Business",
  589107:"Run the Business",589108:"Run the Business",589110:"Run the Business",589112:"Run the Business",
  589122:"Unplanned / Ad-hoc Work",589142:"Run the Business",589145:"Planned Delivery",589150:"Planned Delivery",
  589162:"Run the Business",589166:"Planned Delivery",589196:"Run the Business",589200:"Overhead",
  589202:"Planned Delivery",589203:"Unplanned / Ad-hoc Work",589205:"Planned Delivery",589206:"Run the Business",
  589208:"Run the Business",589257:"Run the Business",589267:"Run the Business",589282:"Planned Delivery",
  589285:"Planned Delivery",589544:"Planned Delivery",589545:"Run the Business",589546:"Run the Business",
  589577:"Planned Delivery",589584:"Planned Delivery",589585:"Run the Business",589589:"Planned Delivery",
  589592:"Planned Delivery",589597:"Planned Delivery",589602:"Overhead",589603:"Overhead",589604:"Overhead",
  589605:"Planned Delivery",589606:"Overhead",589607:"Overhead",589609:"Planned Delivery",589622:"Planned Delivery",
  589624:"Run the Business",589625:"Run the Business",589626:"Planned Delivery",589627:"Planned Delivery",
  589628:"Planned Delivery",589629:"Run the Business",589631:"Planned Delivery",589638:"Run the Business",
  589663:"Unplanned / Ad-hoc Work",589687:"Planned Delivery",589703:"Overhead",589705:"Planned Delivery",
  589708:"Overhead",590142:"Unplanned / Ad-hoc Work",590497:"Unplanned / Ad-hoc Work",
};
const RBT = {
  568362:"Build",569088:"Transform",585785:"Run",586680:"Build",588673:"Build",589085:"Build",589105:"Run",
  589106:"Run",589107:"Run",589108:"Run",589110:"Run",589112:"Run",589122:"Run",589142:"Run",589145:"Transform",
  589150:"Build",589162:"Run",589166:"Build",589196:"Run",589200:"Run",589202:"Build",589203:"Build",589205:"Build",
  589206:"Run",589208:"Run",589257:"Run",589267:"Run",589282:"Build",589285:"Build",589544:"Build",589545:"Run",
  589546:"Run",589577:"Transform",589584:"Build",589585:"Run",589589:"Build",589592:"Build",589597:"Build",
  589602:"Run",589603:"Run",589604:"Run",589605:"Build",589606:"Run",589607:"Run",589609:"Build",589622:"Transform",
  589624:"Run",589625:"Run",589626:"Build",589627:"Build",589628:"Build",589629:"Run",589631:"Transform",
  589638:"Run",589663:"Run",589687:"Transform",589703:"Run",589705:"Transform",589708:"Run",590142:"Run",590497:"Build",
};
const SPILL_NOTE = {
  568362:"In progress - renewal signals on Engagement Signals dashboard; continues next sprint.",
  569088:"Not started - AI-enabled Signals replacement scheduled behind the renewal-signals work.",
  589166:"November MBU deck template - month-start deliverable; continues into Week 44.",
};

// patch orphan MBU November feature -> Sales Excellence
for (const s of STORIES) {
  if (s.initiative === "Unmapped") s.initiative = "Sales Excellence";
  s.category = CATEGORY[s.id] || "";
  s.lens = RBT[s.id] || "";
  s.closed = ["Closed", "Resolved"].includes(s.state);
  s.completed = s.closed ? s.est : 0;
}
const SA = STORIES.filter(s => s.sprintAgreed), MW = STORIES.filter(s => !s.sprintAgreed);
const CLOSED = STORIES.filter(s => s.closed), SPILL = STORIES.filter(s => !s.closed);
const sum = (a, k = "est") => Math.round(a.reduce((x, s) => x + s[k], 0) * 10) / 10;
const BASE_CAP = 312;
const ABSENCES = 36; // Juan Guerrero full week (12) + Malena 2 days (16) + Juanse Uribe Monday (8)
const CAP = BASE_CAP - ABSENCES; // 276 adjusted
const stats = {
  saN: SA.length, saH: sum(SA), saClosed: SA.filter(s => s.closed).length,
  mwN: MW.length, mwH: sum(MW), mwClosed: MW.filter(s => s.closed).length,
  clN: CLOSED.length, clH: sum(CLOSED),
  spN: SPILL.length, spH: sum(SPILL),
};
console.log(JSON.stringify(stats));

// impact matrix
const INITS = ["In-Year Parenting", "Sales Excellence", "Reporting & Automation", "Employee Giving"];
const matrix = INITS.map(init => {
  const g = CLOSED.filter(s => s.initiative === init);
  const r = sum(g.filter(s => s.lens === "Run")), b = sum(g.filter(s => s.lens === "Build")), t = sum(g.filter(s => s.lens === "Transform"));
  const tot = Math.round((r + b + t) * 10) / 10;
  return [init, r, b, t, tot, tot ? Math.round((b + t) / tot * 1000) / 10 + "%" : "0%"];
});
const TR = sum(CLOSED.filter(s => s.lens === "Run")), TB = sum(CLOSED.filter(s => s.lens === "Build")), TT = sum(CLOSED.filter(s => s.lens === "Transform"));
matrix.push(["TOTAL", TR, TB, TT, stats.clH, Math.round((TB + TT) / stats.clH * 1000) / 10 + "%"]);
console.log(JSON.stringify(matrix));

// category summary (all committed)
const CATS = ["Planned Delivery", "Run the Business", "Overhead", "Unplanned / Ad-hoc Work"];
const catRows = INITS.map(init => {
  const g = STORIES.filter(s => s.initiative === init);
  const row = [init];
  for (const c of CATS) { const gg = g.filter(s => s.category === c); row.push(gg.length, sum(gg)); }
  row.push(g.length, sum(g));
  return row;
});

// ---------- palette & helpers (same design language as before) ----------
const NAVY = "FF1F3864", BLUE = "FF2E75B6", BAND = "FFF2F7FC", LINE = "FFB7CCE4";
function solid(argb) { return { type: "pattern", pattern: "solid", fgColor: { argb } }; }
function thin(argb = LINE) { return { style: "thin", color: { argb } }; }
const STATUS = {
  "Closed": { fill: "FFC6EFCE", font: "FF006100" }, "Resolved": { fill: "FFC6EFCE", font: "FF006100" },
  "In Progress": { fill: "FFFFF2CC", font: "FF7F6000" }, "Active": { fill: "FFFFF2CC", font: "FF7F6000" },
  "Not Started": { fill: "FFF2F2F2", font: "FF595959" }, "New": { fill: "FFF2F2F2", font: "FF595959" },
  "On Hold": { fill: "FFFFC7CE", font: "FF9C0006" }, "SpillOver": { fill: "FFFFC7CE", font: "FF9C0006" },
};
const CATP = {
  "Planned Delivery": { fill: "FFDDEBF7", font: "FF1F4E79" }, "Run the Business": { fill: "FFE2EFDA", font: "FF375623" },
  "Overhead": { fill: "FFEDEDED", font: "FF595959" }, "Unplanned / Ad-hoc Work": { fill: "FFFCE4D6", font: "FF833C00" },
};
const LENSP = {
  "Run": { fill: "FFDEEBF7", font: "FF1F4E79" }, "Build": { fill: "FFE2EFDA", font: "FF375623" }, "Transform": { fill: "FFE4DFEC", font: "FF403152" },
};
const INITP = {
  "In-Year Parenting": { fill: "FFFFE6CC", font: "FF7F4F00" }, "Sales Excellence": { fill: "FFEADCF5", font: "FF5B2D81" },
  "Reporting & Automation": { fill: "FFD6E6F7", font: "FF1F4E79" }, "Employee Giving": { fill: "FFDDF0DD", font: "FF2E6B2E" },
};

const wb = new ExcelJS.Workbook();
wb.creator = "Antonio Alvarez"; wb.created = new Date();

function banner(ws, title, subtitle, lastCol) {
  ws.mergeCells(1, 1, 1, lastCol); ws.mergeCells(2, 1, 2, lastCol);
  ws.getRow(1).height = 32; ws.getRow(2).height = 18;
  const t = ws.getCell(1, 1); t.value = title;
  t.font = { name: "Segoe UI", size: 18, bold: true, color: { argb: "FFFFFFFF" } };
  t.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  const s = ws.getCell(2, 1); s.value = subtitle;
  s.font = { name: "Segoe UI", size: 10, italic: true, color: { argb: "FFD6E4F0" } };
  s.alignment = { vertical: "top", horizontal: "left", indent: 1 };
  for (let c = 1; c <= lastCol; c++) { ws.getCell(1, c).fill = solid(NAVY); ws.getCell(2, c).fill = solid(NAVY); }
  ws.getRow(3).height = 6;
}
function styleHeader(ws, row, nCols) {
  ws.getRow(row).height = 30;
  for (let c = 1; c <= nCols; c++) {
    const cell = ws.getCell(row, c);
    cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = solid(BLUE);
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = { top: thin(), bottom: thin(), left: thin("FFFFFFFF"), right: thin("FFFFFFFF") };
  }
}
function pill(cell, map, val) {
  const s = map[val]; if (!s) return;
  cell.fill = solid(s.fill);
  cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: s.font } };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}
function dataTable(ws, startRow, headers, rows, opts) {
  const o = Object.assign({ linkCols: [], statusCols: [], catCols: [], lensCols: [], initCols: [], numCols: [], wrapCols: [] }, opts);
  headers.forEach((h, i) => { ws.getCell(startRow, i + 1).value = h; });
  styleHeader(ws, startRow, headers.length);
  rows.forEach((rowVals, ri) => {
    const r = startRow + 1 + ri, banded = ri % 2 === 1;
    rowVals.forEach((val, ci) => {
      const cell = ws.getCell(r, ci + 1);
      if (o.linkCols.includes(ci) && val != null && val !== "") {
        const id = String(val);
        cell.value = { text: "#" + id, hyperlink: ADO(id), tooltip: "Open work item " + id + " in Azure DevOps" };
        cell.font = { name: "Segoe UI", size: 10, color: { argb: "FF0563C1" }, underline: true };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      } else {
        cell.value = val;
        cell.font = { name: "Segoe UI", size: 10, color: { argb: "FF333333" } };
        cell.alignment = { horizontal: o.numCols.includes(ci) ? "center" : "left", vertical: "middle", wrapText: o.wrapCols.includes(ci) };
      }
      if (banded) cell.fill = solid(BAND);
      cell.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
      if (o.statusCols.includes(ci)) pill(cell, STATUS, val);
      if (o.catCols.includes(ci)) pill(cell, CATP, val);
      if (o.lensCols.includes(ci)) pill(cell, LENSP, val);
      if (o.initCols.includes(ci)) pill(cell, INITP, val);
      if (o.numCols.includes(ci) && typeof val === "number") cell.numFmt = "0.0#";
    });
    ws.getRow(r).height = o.rowHeight || 18;
  });
  if (o.widths) o.widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  ws.autoFilter = { from: { row: startRow, column: 1 }, to: { row: startRow + rows.length, column: headers.length } };
  ws.views = [{ state: "frozen", ySplit: startRow }];
}
function sectionLabel(ws, row, col, text, span = 3) {
  ws.mergeCells(row, col, row, col + span - 1);
  const c = ws.getCell(row, col); c.value = text;
  c.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: NAVY } };
  c.alignment = { vertical: "middle" };
  ws.getRow(row).height = 22;
}
function statStrip(ws, items) {
  let col = 1; const row = 4;
  ws.getRow(row).height = 24;
  for (const s of items) {
    ws.mergeCells(row, col, row, col + 1);
    const l = ws.getCell(row, col); l.value = s.label;
    l.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF595959" } };
    l.alignment = { horizontal: "right", vertical: "middle" };
    const v = ws.getCell(row, col + 2); v.value = s.value;
    v.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: BLUE } };
    v.alignment = { horizontal: "center", vertical: "middle" };
    v.fill = solid("FFEFF5FB");
    v.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
    col += 4;
  }
  ws.getRow(5).height = 8;
}
function kpiCards(ws, row, cards, perCard = 2) {
  let col = 1;
  ws.getRow(row).height = 26; ws.getRow(row + 1).height = 16;
  const e = { style: "medium", color: { argb: LINE } };
  for (const card of cards) {
    ws.mergeCells(row, col, row, col + perCard - 1);
    ws.mergeCells(row + 1, col, row + 1, col + perCard - 1);
    const v = ws.getCell(row, col); v.value = card.value;
    v.font = { name: "Segoe UI", size: 16, bold: true, color: { argb: card.accent || BLUE } };
    v.alignment = { horizontal: "center", vertical: "middle" };
    const l = ws.getCell(row + 1, col); l.value = card.label;
    l.font = { name: "Segoe UI", size: 9, color: { argb: "FF595959" } };
    l.alignment = { horizontal: "center", vertical: "top", wrapText: true };
    for (let c = col; c < col + perCard; c++) {
      ws.getCell(row, c).border = { top: e, left: c === col ? e : undefined, right: c === col + perCard - 1 ? e : undefined };
      ws.getCell(row + 1, c).border = { bottom: e, left: c === col ? e : undefined, right: c === col + perCard - 1 ? e : undefined };
    }
    col += perCard;
  }
}

// ============ Dashboard ============
{
  const ws = wb.addWorksheet("Dashboard", { properties: { tabColor: { argb: NAVY } }, views: [{ showGridLines: false }] });
  banner(ws, "Week 43 \u2014 Delivery & Impact Summary", "ADO Sprint Week 43 - 2026 \u00b7 Oct 26 \u2013 Oct 30, 2026 \u00b7 TSI Sales and Ops \u00b7 generated from Azure DevOps", 12);
  kpiCards(ws, 5, [
    { label: "Adjusted Capacity (hrs)", value: CAP, accent: "FF595959" },
    { label: "Committed (hrs)", value: stats.saH + stats.mwH, accent: BLUE },
    { label: "Completed Value (hrs)", value: stats.clH, accent: "FF2E7D32" },
    { label: "Sprint-Agreed Closure", value: Math.round(stats.saClosed / stats.saN * 100) + "%", accent: "FF2E7D32" },
    { label: "Mid-Week Closure", value: "100%", accent: "FF2E7D32" },
    { label: "Innovation Ratio", value: Math.round((TB + TT) / stats.clH * 100) + "%", accent: "FF6A329F" },
  ], 2);
  ws.getRow(7).height = 10;

  sectionLabel(ws, 8, 1, "Impact Analysis \u2014 closed / resolved hours by initiative", 6);
  const imHdr = ["Initiative", "\u{1F501} Run (hrs)", "\u{1F6E0} Build (hrs)", "\u{1F680} Transform (hrs)", "Total (hrs)", "Innovation Ratio"];
  imHdr.forEach((h, i) => ws.getCell(9, i + 1).value = h);
  styleHeader(ws, 9, 6);
  matrix.forEach((rv, ri) => {
    const r = 10 + ri, isTotal = rv[0] === "TOTAL";
    rv.forEach((v, ci) => {
      const cell = ws.getCell(r, ci + 1); cell.value = v;
      cell.font = { name: "Segoe UI", size: 10.5, bold: isTotal || ci === 0, color: { argb: isTotal ? NAVY : "FF333333" } };
      cell.alignment = { horizontal: ci === 0 ? "left" : "center", vertical: "middle", indent: ci === 0 ? 1 : 0 };
      cell.fill = solid(isTotal ? "FFD6E4F0" : ri % 2 ? BAND : "FFFFFFFF");
      cell.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
      if (typeof v === "number") cell.numFmt = "0.0";
    });
    if (!isTotal) pill(ws.getCell(r, 1), INITP, rv[0]);
    ws.getRow(r).height = 20;
  });

  sectionLabel(ws, 16, 1, "Impact lens definitions", 6);
  const defs = [
    ["Run", "Work that keeps existing reports, pipelines and operations running (refreshes, support, governance, ad-hoc fixes)."],
    ["Build", "Enhancements and new deliverables on existing assets (new views, measures, forms, validations that add capability)."],
    ["Transform", "New capabilities or architecture change (migrations, Dev-to-Prod promotion, AI tooling, new data platforms)."],
  ];
  defs.forEach((d, i) => {
    const r = 17 + i;
    const k = ws.getCell(r, 1); k.value = d[0]; pill(k, LENSP, d[0]);
    k.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
    ws.mergeCells(r, 2, r, 6);
    const v = ws.getCell(r, 2); v.value = d[1];
    v.font = { name: "Segoe UI", size: 10, color: { argb: "FF444444" } };
    v.alignment = { vertical: "middle", wrapText: true, indent: 1 };
    v.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
    ws.getRow(r).height = 26;
  });

  sectionLabel(ws, 21, 1, "User story category mix \u2014 all committed stories (count / hrs)", 11);
  const catHdr = ["Initiative", "Planned Delivery #", "Planned Delivery hrs", "Run the Business #", "Run the Business hrs", "Overhead #", "Overhead hrs", "Ad-hoc #", "Ad-hoc hrs", "Total #", "Total hrs"];
  catHdr.forEach((h, i) => ws.getCell(22, i + 1).value = h);
  styleHeader(ws, 22, 11);
  const tot = ["TOTAL"]; for (let i = 1; i < 11; i++) tot.push(Math.round(catRows.reduce((s, r) => s + (r[i] || 0), 0) * 10) / 10);
  [...catRows, tot].forEach((rv, ri) => {
    const r = 23 + ri, isTotal = rv[0] === "TOTAL";
    rv.forEach((v, ci) => {
      const cell = ws.getCell(r, ci + 1); cell.value = v;
      cell.font = { name: "Segoe UI", size: 10, bold: isTotal || ci === 0, color: { argb: isTotal ? NAVY : "FF333333" } };
      cell.alignment = { horizontal: ci === 0 ? "left" : "center", vertical: "middle", indent: ci === 0 ? 1 : 0 };
      cell.fill = solid(isTotal ? "FFD6E4F0" : ri % 2 ? BAND : "FFFFFFFF");
      cell.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
      if (typeof v === "number" && String(catHdr[ci]).includes("hrs")) cell.numFmt = "0.0";
    });
    if (!isTotal) pill(ws.getCell(r, 1), INITP, rv[0]);
    ws.getRow(r).height = 19;
  });

  sectionLabel(ws, 29, 1, "Notes", 11);
  const notes = [
    "Sprint-agreed = the 59 stories created in / moved into the sprint through end of Tuesday 10/27 (planning ran Mon-Tue). Mid-week = 2 stories raised 10/28-10/29.",
    "Capacity adjusted for availability: standard 312 hrs minus 36 hrs of absences - Juan Guerrero out the full week, sick (12 hrs, 0.3 FTE); Malena Arduino out 2 days (16 hrs); Juanse Uribe out Monday (8 hrs). Adjusted capacity = 276 hrs.",
    "Absences did not drive spillover: Juan Guerrero had no committed stories; Malena closed all 7 of her stories (23 hrs) in 3 working days; Juanse closed all 6 of his (25 hrs). The 3 open stories belong to a different workstream.",
    "Impact matrix counts closed/resolved hours only (" + stats.clH + "). 3 open stories (13 hrs) carry into next sprint.",
  ];
  notes.forEach((n, i) => {
    const r = 30 + i;
    ws.mergeCells(r, 1, r, 11);
    const c = ws.getCell(r, 1); c.value = "\u2022 " + n;
    c.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF595959" } };
    c.alignment = { wrapText: true, vertical: "middle" };
    ws.getRow(r).height = 24;
  });
  for (let i = 1; i <= 12; i++) ws.getColumn(i).width = 13.5;
  ws.getColumn(1).width = 24;
}

// ============ Sprint Agreed ============
{
  const hdr = ["ID", "Title", "Assigned To", "Created", "State (end of week)", "Original Estimate", "Remaining", "Completed", "User Story Category", "Initiative", "End-of-Week Status"];
  const rows = SA.map(s => [s.id, s.title, s.assigned, s.created, s.state, s.est, s.closed ? 0 : s.rem, s.completed, s.category, s.initiative, s.closed ? "Closed" : "SpillOver"]);
  const ws = wb.addWorksheet("Sprint Agreed Items", { properties: { tabColor: { argb: BLUE } }, views: [{ showGridLines: false }] });
  banner(ws, `Sprint Agreed Items \u2014 ${stats.saN} stories \u00b7 ${stats.saH} hrs`, `In the sprint by end of Tuesday 10/27 \u00b7 ${Math.round(stats.saH / CAP * 100)}% of ${CAP}-hr capacity \u00b7 ${stats.saClosed} closed (${Math.round(stats.saClosed / stats.saN * 100)}%)`, hdr.length);
  statStrip(ws, [
    { label: "Adjusted Capacity", value: CAP },
    { label: "Original Estimate", value: stats.saH },
    { label: "Capacity Estimation", value: Math.round(stats.saH / CAP * 1000) / 10 + "%" },
  ]);
  dataTable(ws, 6, hdr, rows, { widths: [9, 58, 24, 12, 15, 11, 11, 11, 20, 22, 14], linkCols: [0], statusCols: [4, 10], catCols: [8], initCols: [9], numCols: [5, 6, 7], wrapCols: [1], rowHeight: 20 });
}

// ============ Mid Week ============
{
  const hdr = ["ID", "Title", "Assigned To", "Created", "State (end of week)", "Original Estimate", "Completed", "User Story Category", "Initiative", "End-of-Week Status"];
  const rows = MW.map(s => [s.id, s.title, s.assigned, s.created, s.state, s.est, s.completed, s.category, s.initiative, s.closed ? "Closed" : "SpillOver"]);
  const ws = wb.addWorksheet("Mid Week Items", { properties: { tabColor: { argb: "FFED7D31" } }, views: [{ showGridLines: false }] });
  banner(ws, `Mid-Week Items \u2014 ${stats.mwN} stories \u00b7 ${stats.mwH} hrs`, `Raised after Tuesday planning \u00b7 ${Math.round(stats.mwH / CAP * 1000) / 10}% of capacity \u00b7 both closed within the week (100%)`, hdr.length);
  statStrip(ws, [
    { label: "Adjusted Capacity", value: CAP },
    { label: "Original Estimate", value: stats.mwH },
    { label: "Capacity Ratio", value: Math.round(stats.mwH / CAP * 1000) / 10 + "%" },
  ]);
  dataTable(ws, 6, hdr, rows, { widths: [9, 58, 24, 12, 15, 11, 11, 20, 22, 14], linkCols: [0], statusCols: [4, 9], catCols: [7], initCols: [8], numCols: [5, 6], wrapCols: [1], rowHeight: 20 });
}

// ============ Closed & Resolved ============
{
  const hdr = ["ID", "Title", "Original Estimate", "Completed", "State", "Item Creation", "Initiative"];
  const rows = CLOSED.map(s => [s.id, s.title, s.est, s.completed, s.state, s.sprintAgreed ? "Sprint Agreed" : "Mid-Week", s.initiative]);
  const ws = wb.addWorksheet("Closed & Resolved Items", { properties: { tabColor: { argb: "FF548235" } }, views: [{ showGridLines: false }] });
  banner(ws, `Closed & Resolved \u2014 ${stats.clN} stories \u00b7 ${stats.clH} hrs completed value`, `Sprint-agreed closure ${Math.round(stats.saClosed / stats.saN * 100)}% \u00b7 Mid-week closure 100%`, hdr.length);
  statStrip(ws, [
    { label: "Sprint Agreed Closure", value: Math.round(stats.saClosed / stats.saN * 1000) / 10 + "%" },
    { label: "Mid Week Closure", value: "100%" },
    { label: "Completed Value", value: stats.clH },
  ]);
  dataTable(ws, 6, hdr, rows, { widths: [9, 66, 12, 12, 12, 14, 22], linkCols: [0], statusCols: [4], initCols: [6], numCols: [2, 3], wrapCols: [1], rowHeight: 20 });
}

// ============ SpillOver ============
{
  const hdr = ["ID", "Title", "State", "Original Estimate", "SpillOver Reason / Note"];
  const rows = SPILL.map(s => [s.id, s.title, s.state, s.est, SPILL_NOTE[s.id] || ""]);
  const ws = wb.addWorksheet("SpillOver Items", { properties: { tabColor: { argb: "FFC00000" } }, views: [{ showGridLines: false }] });
  banner(ws, `SpillOver \u2014 ${stats.spN} stories \u00b7 ${stats.spH} hrs`, "Open at end of week; all three carry into the next sprint", hdr.length);
  statStrip(ws, [
    { label: "Total Items", value: stats.spN },
    { label: "Original Estimate", value: stats.spH },
  ]);
  dataTable(ws, 6, hdr, rows, { widths: [9, 62, 14, 12, 62], linkCols: [0], statusCols: [2], numCols: [3], wrapCols: [1, 4], rowHeight: 26 });
}

// ============ Friday's Email Stories ============
{
  const hdr = ["ID", "Title", "Assigned To", "State", "Tags", "Describe Business Impact", "Parent", "Original Estimate", "Completed", "Item Creation", "User Story Category", "Impact Lens", "Initiative"];
  const rows = STORIES.map(s => [s.id, s.title, s.assigned, s.state, s.tags, s.impact, s.parent, s.est, s.completed, s.sprintAgreed ? "Sprint Agreed" : "Mid-Week", s.category, s.lens, s.initiative]);
  const ws = wb.addWorksheet("Friday's Email Stories", { properties: { tabColor: { argb: "FF6A329F" } }, views: [{ showGridLines: false }] });
  banner(ws, `Friday's Email Stories \u2014 all ${STORIES.length} week-43 stories`, "Business impact text per story \u00b7 source for the Impact Analysis and Scorecard", hdr.length);
  ws.getRow(4).height = 4; ws.getRow(5).height = 4;
  dataTable(ws, 6, hdr, rows, { widths: [9, 42, 22, 12, 16, 70, 10, 10, 10, 13, 20, 12, 22], linkCols: [0, 6], statusCols: [3], catCols: [10], lensCols: [11], initCols: [12], numCols: [7, 8], wrapCols: [1, 4, 5], rowHeight: 42 });
}

// ============ Features to Epic ============
{
  const hdr = ["Feature ID", "Feature Title", "Epic ID", "Epic Name", "Initiative ID", "Initiative Name"];
  const rows = F2E.slice(1).filter(r => r && r[0] != null);
  const ws = wb.addWorksheet("Features to Epic", { properties: { tabColor: { argb: "FF808080" } }, views: [{ showGridLines: false }] });
  banner(ws, "Features \u2192 Epic \u2192 Initiative", "Parent-chain mapping used to assign each story to an initiative (new: MBU November \u2192 Sales Excellence)", hdr.length);
  ws.getRow(4).height = 4; ws.getRow(5).height = 4;
  dataTable(ws, 6, hdr, rows, { widths: [12, 52, 12, 34, 13, 40], linkCols: [0, 2, 4], wrapCols: [1, 3, 5], rowHeight: 18 });
}

// ============ Planning Baseline ============
{
  const hdr = ["ID", "Title", "Created", "Original Estimate", "User Story Category", "Impact Lens", "Initiative"];
  const rows = SA.map(s => [s.id, s.title, s.created, s.est, s.category, s.lens, s.initiative]);
  const ws = wb.addWorksheet("Planning Baseline", { properties: { tabColor: { argb: "FF808080" } }, views: [{ showGridLines: false }] });
  banner(ws, "Planning Baseline \u2014 the 59-story sprint plan as of Tue 10/27", "Derived from ADO creation / iteration-entry timestamps (planning ran Mon 10/26 \u2013 Tue 10/27)", hdr.length);
  ws.getRow(4).height = 4; ws.getRow(5).height = 4;
  dataTable(ws, 6, hdr, rows, { widths: [9, 62, 12, 12, 22, 12, 24], linkCols: [0], catCols: [4], lensCols: [5], initCols: [6], wrapCols: [1], rowHeight: 20 });
}

const out = require("path").join(OUT, "103026 Friday Email.xlsx");
wb.xlsx.writeFile(out).then(() => console.log("WROTE", out));
