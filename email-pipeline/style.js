// Builds a styled Week 43 Friday Email workbook from data.json using exceljs
const ExcelJS = require("exceljs");
const fs = require("fs");
const D = JSON.parse(fs.readFileSync("data.json"));
const OUT = process.argv[2] || ".";

const ADO = id => `https://dev.azure.com/oneCELA/TSI%20Sales%20and%20Ops/_workitems/edit/${id}`;

// ---------- palette ----------
const NAVY = "FF1F3864", BLUE = "FF2E75B6", BAND = "FFF2F7FC", LINE = "FFB7CCE4";
const C = {
  title: { font: { name: "Segoe UI", size: 18, bold: true, color: { argb: "FFFFFFFF" } }, fill: solid(NAVY), alignment: { vertical: "middle", horizontal: "left", indent: 1 } },
  sub:   { font: { name: "Segoe UI", size: 10, italic: true, color: { argb: "FFD6E4F0" } }, fill: solid(NAVY), alignment: { vertical: "top", horizontal: "left", indent: 1 } },
};
function solid(argb) { return { type: "pattern", pattern: "solid", fgColor: { argb } }; }

const STATUS = {
  "Closed":      { fill: "FFC6EFCE", font: "FF006100" },
  "Resolved":    { fill: "FFC6EFCE", font: "FF006100" },
  "In Progress": { fill: "FFFFF2CC", font: "FF7F6000" },
  "Active":      { fill: "FFFFF2CC", font: "FF7F6000" },
  "Not Started": { fill: "FFF2F2F2", font: "FF595959" },
  "New":         { fill: "FFF2F2F2", font: "FF595959" },
  "On Hold":     { fill: "FFFFC7CE", font: "FF9C0006" },
  "SpillOver":   { fill: "FFFFC7CE", font: "FF9C0006" },
  "Moved to next sprint": { fill: "FFFFC7CE", font: "FF9C0006" },
};
const CAT = {
  "Planned Delivery":        { fill: "FFDDEBF7", font: "FF1F4E79" },
  "Run the Business":        { fill: "FFE2EFDA", font: "FF375623" },
  "Overhead":                { fill: "FFEDEDED", font: "FF595959" },
  "Unplanned / Ad-hoc Work": { fill: "FFFCE4D6", font: "FF833C00" },
};
const LENS = {
  "Run":       { fill: "FFDEEBF7", font: "FF1F4E79" },
  "Build":     { fill: "FFE2EFDA", font: "FF375623" },
  "Transform": { fill: "FFE4DFEC", font: "FF403152" },
};
const INIT = {
  "In-Year Parenting":      { fill: "FFFFE6CC", font: "FF7F4F00" },
  "Sales Excellence":       { fill: "FFEADCF5", font: "FF5B2D81" },
  "Reporting & Automation": { fill: "FFD6E6F7", font: "FF1F4E79" },
  "Employee Giving":        { fill: "FFDDF0DD", font: "FF2E6B2E" },
};

const wb = new ExcelJS.Workbook();
wb.creator = "Antonio Alvarez";
wb.created = new Date();

// ---------- helpers ----------
function banner(ws, title, subtitle, lastCol) {
  ws.mergeCells(1, 1, 1, lastCol);
  ws.mergeCells(2, 1, 2, lastCol);
  ws.getRow(1).height = 32; ws.getRow(2).height = 18;
  const t = ws.getCell(1, 1); t.value = title; Object.assign(t, { font: C.title.font, fill: C.title.fill, alignment: C.title.alignment });
  const s = ws.getCell(2, 1); s.value = subtitle; Object.assign(s, { font: C.sub.font, fill: C.sub.fill, alignment: C.sub.alignment });
  for (let c = 1; c <= lastCol; c++) { ws.getCell(1, c).fill = solid(NAVY); ws.getCell(2, c).fill = solid(NAVY); }
  ws.getRow(3).height = 6;
}

function kpiCards(ws, row, cards, perCard = 2) {
  // cards: [{label, value, accent}]
  let col = 1;
  ws.getRow(row).height = 26; ws.getRow(row + 1).height = 16;
  for (const card of cards) {
    ws.mergeCells(row, col, row, col + perCard - 1);
    ws.mergeCells(row + 1, col, row + 1, col + perCard - 1);
    const v = ws.getCell(row, col);
    v.value = card.value;
    v.font = { name: "Segoe UI", size: 16, bold: true, color: { argb: card.accent || BLUE } };
    v.alignment = { horizontal: "center", vertical: "middle" };
    v.fill = solid("FFFFFFFF");
    const l = ws.getCell(row + 1, col);
    l.value = card.label;
    l.font = { name: "Segoe UI", size: 9, color: { argb: "FF595959" } };
    l.alignment = { horizontal: "center", vertical: "top", wrapText: true };
    for (let c = col; c < col + perCard; c++) {
      for (let r = row; r <= row + 1; r++) {
        ws.getCell(r, c).border = boxBorder(r === row, r === row + 1, c === col, c === col + perCard - 1);
      }
    }
    col += perCard;
  }
}
function boxBorder(top, bottom, left, right) {
  const e = { style: "medium", color: { argb: LINE } };
  return { top: top ? e : undefined, bottom: bottom ? e : undefined, left: left ? e : undefined, right: right ? e : undefined };
}

function styleHeader(ws, row, nCols) {
  const r = ws.getRow(row);
  r.height = 30;
  for (let c = 1; c <= nCols; c++) {
    const cell = ws.getCell(row, c);
    cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = solid(BLUE);
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = { top: thin(), bottom: thin(), left: thin("FFFFFFFF"), right: thin("FFFFFFFF") };
  }
}
function thin(argb = LINE) { return { style: "thin", color: { argb } }; }

function pill(cell, map, val) {
  const s = map[val];
  if (!s) return;
  cell.fill = solid(s.fill);
  cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: s.font } };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}

function dataTable(ws, startRow, headers, rows, opts) {
  // opts: widths[], linkCols:[idx], statusCols:[idx], catCols:[idx], lensCols:[idx], initCols:[idx], numCols:[idx], wrapCols:[idx], dateCols:[idx]
  const o = Object.assign({ linkCols: [], statusCols: [], catCols: [], lensCols: [], initCols: [], numCols: [], wrapCols: [], dateCols: [] }, opts);
  headers.forEach((h, i) => { ws.getCell(startRow, i + 1).value = h; });
  styleHeader(ws, startRow, headers.length);
  rows.forEach((rowVals, ri) => {
    const r = startRow + 1 + ri;
    const banded = ri % 2 === 1;
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
      if (o.catCols.includes(ci)) pill(cell, CAT, val);
      if (o.lensCols.includes(ci)) pill(cell, LENS, val);
      if (o.initCols.includes(ci)) pill(cell, INIT, val);
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
  const c = ws.getCell(row, col);
  c.value = text;
  c.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: NAVY } };
  c.alignment = { vertical: "middle" };
  ws.getRow(row).height = 22;
}

// strip old stats block + header row; return data rows (drop ADO Link col at index 1)
function tableRows(sheet, headerRowIdx, dropCols = [1]) {
  return D[sheet].slice(headerRowIdx + 1).filter(r => r && r[0] != null)
    .map(r => r.filter((_, i) => !dropCols.includes(i)));
}

// =======================================================================
// 1. DASHBOARD
// =======================================================================
{
  const ws = wb.addWorksheet("Dashboard", { properties: { tabColor: { argb: NAVY } }, views: [{ showGridLines: false }] });
  const NC = 12;
  banner(ws, "Week 43 \u2014 Delivery & Impact Summary", "Elevate Sprint \u00b7 Oct 19 \u2013 Oct 23, 2026 \u00b7 TSI Sales and Ops \u00b7 generated from Azure DevOps", NC);

  kpiCards(ws, 5, [
    { label: "Team Capacity (hrs)", value: 312, accent: "FF595959" },
    { label: "Committed (hrs)", value: 293.7, accent: BLUE },
    { label: "Completed Value (hrs)", value: 259.2, accent: "FF2E7D32" },
    { label: "Sprint-Agreed Closure", value: "83%", accent: "FF2E7D32" },
    { label: "Mid-Week Closure", value: "90%", accent: "FF2E7D32" },
    { label: "Innovation Ratio", value: "52%", accent: "FF6A329F" },
  ], 2);
  ws.getRow(7).height = 10;

  // Impact matrix
  sectionLabel(ws, 8, 1, "Impact Analysis \u2014 closed / resolved hours by initiative", 6);
  const imHdr = ["Initiative", "\u{1F501} Run (hrs)", "\u{1F6E0} Build (hrs)", "\u{1F680} Transform (hrs)", "Total (hrs)", "Innovation Ratio"];
  const imRows = D["Impact Summary"].slice(3, 8).map(r => r.slice(0, 6));
  imHdr.forEach((h, i) => ws.getCell(9, i + 1).value = h);
  styleHeader(ws, 9, 6);
  imRows.forEach((rv, ri) => {
    const r = 10 + ri, isTotal = rv[0] === "TOTAL";
    rv.forEach((v, ci) => {
      const cell = ws.getCell(r, ci + 1);
      cell.value = v;
      cell.font = { name: "Segoe UI", size: 10.5, bold: isTotal || ci === 0, color: { argb: isTotal ? NAVY : "FF333333" } };
      cell.alignment = { horizontal: ci === 0 ? "left" : "center", vertical: "middle", indent: ci === 0 ? 1 : 0 };
      cell.fill = solid(isTotal ? "FFD6E4F0" : ri % 2 ? BAND : "FFFFFFFF");
      cell.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
      if (typeof v === "number") cell.numFmt = "0.0";
    });
    if (!isTotal) pill(ws.getCell(r, 1), INIT, rv[0]);
    ws.getRow(r).height = 20;
  });

  // definitions
  sectionLabel(ws, 16, 1, "Impact lens definitions", 6);
  const defs = D["Impact Summary"].slice(10, 13);
  defs.forEach((d, i) => {
    const r = 17 + i;
    const k = ws.getCell(r, 1); k.value = d[0];
    pill(k, LENS, d[0]);
    k.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
    ws.mergeCells(r, 2, r, 6);
    const v = ws.getCell(r, 2); v.value = d[1];
    v.font = { name: "Segoe UI", size: 10, color: { argb: "FF444444" } };
    v.alignment = { vertical: "middle", wrapText: true, indent: 1 };
    v.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
    ws.getRow(r).height = 26;
  });

  // category summary
  sectionLabel(ws, 21, 1, "User story category mix \u2014 all committed stories (count / hrs)", 11);
  const catHdr = ["Initiative", "Planned Delivery #", "Planned Delivery hrs", "Run the Business #", "Run the Business hrs", "Overhead #", "Overhead hrs", "Ad-hoc #", "Ad-hoc hrs", "Total #", "Total hrs"];
  catHdr.forEach((h, i) => ws.getCell(22, i + 1).value = h);
  styleHeader(ws, 22, 11);
  const catRows = D["Impact Summary"].slice(16, 20).map(r => r.slice(0, 11));
  const tot = ["TOTAL"]; for (let i = 1; i < 11; i++) tot.push(catRows.reduce((s, r) => s + (r[i] || 0), 0));
  catRows.push(tot);
  catRows.forEach((rv, ri) => {
    const r = 23 + ri, isTotal = rv[0] === "TOTAL";
    rv.forEach((v, ci) => {
      const cell = ws.getCell(r, ci + 1);
      cell.value = v;
      cell.font = { name: "Segoe UI", size: 10, bold: isTotal || ci === 0, color: { argb: isTotal ? NAVY : "FF333333" } };
      cell.alignment = { horizontal: ci === 0 ? "left" : "center", vertical: "middle", indent: ci === 0 ? 1 : 0 };
      cell.fill = solid(isTotal ? "FFD6E4F0" : ri % 2 ? BAND : "FFFFFFFF");
      cell.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
      if (typeof v === "number" && String(catHdr[ci]).includes("hrs")) cell.numFmt = "0.0";
    });
    if (!isTotal) pill(ws.getCell(r, 1), INIT, rv[0]);
    ws.getRow(r).height = 19;
  });

  // notes
  sectionLabel(ws, 29, 1, "Notes", 11);
  const notes = [
    "Sprint-agreed = the 63 stories in Monday 10/19 planning (239.2 hrs, 77% of capacity). Mid-week = 10 stories raised after planning (54.5 hrs, 17%).",
    "3 planned stories were re-planned into the next sprint during the week; 5 weekly leadership-oversight stories are pending administrative closure in ADO (16.5 hrs).",
    "Impact matrix counts closed/resolved hours only (257.7). Completed Value (259.2) adds partial credit on 2 in-progress stories.",
  ];
  notes.forEach((n, i) => {
    const r = 30 + i;
    ws.mergeCells(r, 1, r, 11);
    const c = ws.getCell(r, 1);
    c.value = "\u2022 " + n;
    c.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF595959" } };
    c.alignment = { wrapText: true, vertical: "middle" };
    ws.getRow(r).height = 24;
  });
  for (let i = 1; i <= 12; i++) ws.getColumn(i).width = 13.5;
  ws.getColumn(1).width = 24;
}

// =======================================================================
// generic stat header on list sheets
// =======================================================================
function statStrip(ws, stats, nCols) {
  // stats: [{label, value}]
  let col = 1;
  const row = 4;
  ws.getRow(row).height = 24;
  const per = 3;
  for (const s of stats) {
    ws.mergeCells(row, col, row, col + 1);
    const l = ws.getCell(row, col);
    l.value = s.label;
    l.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF595959" } };
    l.alignment = { horizontal: "right", vertical: "middle" };
    const v = ws.getCell(row, col + 2);
    v.value = s.value;
    v.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: BLUE } };
    v.alignment = { horizontal: "center", vertical: "middle" };
    v.fill = solid("FFEFF5FB");
    v.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
    col += per + 1;
  }
  ws.getRow(5).height = 8;
}

// =======================================================================
// 2. Sprint Agreed Items
// =======================================================================
{
  const hdr = ["ID", "Title", "Assigned To", "State (end of week)", "Original Estimate", "Remaining", "Completed", "User Story Category", "Initiative", "End-of-Week Status"];
  const rows = tableRows("Sprint Agreed Items", 4);
  const ws = wb.addWorksheet("Sprint Agreed Items", { properties: { tabColor: { argb: BLUE } }, views: [{ showGridLines: false }] });
  banner(ws, "Sprint Agreed Items \u2014 63 stories \u00b7 239.2 hrs", "Planned in Monday 10/19 sprint planning \u00b7 77% of 312-hr capacity \u00b7 52 closed (83%)", hdr.length);
  statStrip(ws, [
    { label: "Team Capacity", value: 312 },
    { label: "Original Estimate", value: 239.2 },
    { label: "Capacity Estimation", value: "76.7%" },
  ], hdr.length);
  dataTable(ws, 6, hdr, rows, {
    widths: [9, 58, 24, 15, 11, 11, 11, 20, 22, 14],
    linkCols: [0], statusCols: [3, 9], catCols: [7], initCols: [8], numCols: [4, 5, 6], wrapCols: [1], rowHeight: 20,
  });
}

// =======================================================================
// 3. Mid Week Items
// =======================================================================
{
  const hdr = ["ID", "Title", "Assigned To", "Created", "State (end of week)", "Original Estimate", "Completed", "User Story Category", "Initiative", "End-of-Week Status"];
  const rows = tableRows("Mid Week Items", 4);
  const ws = wb.addWorksheet("Mid Week Items", { properties: { tabColor: { argb: "FFED7D31" } }, views: [{ showGridLines: false }] });
  banner(ws, "Mid-Week Items \u2014 10 stories \u00b7 54.5 hrs", "Raised after Monday planning \u00b7 17% of capacity \u00b7 9 closed within the week (90%)", hdr.length);
  statStrip(ws, [
    { label: "Team Capacity", value: 312 },
    { label: "Original Estimate", value: 54.5 },
    { label: "Capacity Ratio", value: "17.5%" },
  ], hdr.length);
  dataTable(ws, 6, hdr, rows, {
    widths: [9, 58, 24, 12, 15, 11, 11, 20, 22, 14],
    linkCols: [0], statusCols: [4, 9], catCols: [7], initCols: [8], numCols: [5, 6], wrapCols: [1], rowHeight: 20,
  });
}

// =======================================================================
// 4. Closed & Resolved
// =======================================================================
{
  const hdr = ["ID", "Work Item Type", "Title", "Original Estimate", "Completed", "State", "Item Creation", "Initiative"];
  const rows = tableRows("Closed & Resolved Items", 4);
  const ws = wb.addWorksheet("Closed & Resolved Items", { properties: { tabColor: { argb: "FF548235" } }, views: [{ showGridLines: false }] });
  banner(ws, "Closed & Resolved \u2014 61 stories \u00b7 259.2 hrs completed value", "Sprint-agreed closure 83% \u00b7 Mid-week closure 90%", hdr.length);
  statStrip(ws, [
    { label: "Sprint Agreed Closure", value: "82.5%" },
    { label: "Mid Week Closure", value: "90%" },
    { label: "Completed Value", value: 259.2 },
  ], hdr.length);
  dataTable(ws, 6, hdr, rows, {
    widths: [9, 14, 62, 12, 12, 12, 14, 22],
    linkCols: [0], statusCols: [5], initCols: [7], numCols: [3, 4], wrapCols: [2], rowHeight: 20,
  });
}

// =======================================================================
// 5. SpillOver
// =======================================================================
{
  const hdr = ["ID", "Title", "State", "Original Estimate", "SpillOver Reason / Note"];
  const rows = tableRows("SpillOver Items", 3);
  const ws = wb.addWorksheet("SpillOver Items", { properties: { tabColor: { argb: "FFC00000" } }, views: [{ showGridLines: false }] });
  banner(ws, "SpillOver \u2014 12 stories \u00b7 36 hrs", "Open at end of week, re-planned to next sprint, or pending administrative closure", hdr.length);
  statStrip(ws, [
    { label: "Total Items", value: 12 },
    { label: "Original Estimate", value: 36 },
  ], hdr.length);
  dataTable(ws, 6, hdr, rows, {
    widths: [9, 62, 14, 12, 62],
    linkCols: [0], statusCols: [2], numCols: [3], wrapCols: [1, 4], rowHeight: 26,
  });
}

// =======================================================================
// 6. Friday's Email Stories (full detail incl. business impact)
// =======================================================================
{
  const hdr = ["ID", "Type", "Title", "Assigned To", "State", "Tags", "Describe Business Impact", "Parent", "Original Estimate", "Completed", "Mid Week Item", "User Story Category", "Impact Lens", "Initiative"];
  const rows = D["Friday's Email Stories"].slice(1).filter(r => r && r[0] != null).map(r => {
    const x = r.slice(0, 14); // drop old Link col
    x[10] = r[10] == null ? "Mid-Week" : "Sprint Agreed"; // old flag: non-null = sprint agreed
    return x;
  });
  const ws = wb.addWorksheet("Friday's Email Stories", { properties: { tabColor: { argb: "FF6A329F" } }, views: [{ showGridLines: false }] });
  banner(ws, "Friday's Email Stories \u2014 all 70 week-43 stories", "Business impact text per story \u00b7 source for the Impact Analysis and Scorecard", hdr.length);
  ws.getRow(4).height = 4; ws.getRow(5).height = 4;
  dataTable(ws, 6, hdr, rows, {
    widths: [9, 11, 42, 22, 12, 18, 70, 10, 10, 10, 13, 20, 12, 22],
    linkCols: [0, 7], statusCols: [4], catCols: [11], lensCols: [12], initCols: [13], numCols: [8, 9], wrapCols: [2, 5, 6], rowHeight: 42,
  });
}

// =======================================================================
// 7. Features to Epic
// =======================================================================
{
  const hdr = ["Feature ID", "Feature Title", "Epic ID", "Epic Name", "Initiative ID", "Initiative Name"];
  const rows = D["Features to Epic"].slice(1).filter(r => r && r[0] != null);
  const ws = wb.addWorksheet("Features to Epic", { properties: { tabColor: { argb: "FF808080" } }, views: [{ showGridLines: false }] });
  banner(ws, "Features \u2192 Epic \u2192 Initiative", "Parent-chain mapping used to assign each story to an initiative", hdr.length);
  ws.getRow(4).height = 4; ws.getRow(5).height = 4;
  dataTable(ws, 6, hdr, rows, {
    widths: [12, 52, 12, 34, 13, 40],
    linkCols: [0, 2, 4], wrapCols: [1, 3, 5], rowHeight: 18,
  });
}

// =======================================================================
// 8. Monday Items
// =======================================================================
{
  const hdr = ["ID", "Title", "State (at planning)", "Original Estimate", "User Story Category", "Initiative"];
  const rows = D["Monday Items"].slice(1).filter(r => r && r[0] != null).map(r => [r[0], r[2], r[3], r[4], r[5], r[6]]);
  const ws = wb.addWorksheet("Monday Items", { properties: { tabColor: { argb: "FF808080" } }, views: [{ showGridLines: false }] });
  banner(ws, "Monday Items \u2014 the 63-story sprint plan as of 10/19", "Baseline used to split Sprint-Agreed vs Mid-Week", hdr.length);
  ws.getRow(4).height = 4; ws.getRow(5).height = 4;
  dataTable(ws, 6, hdr, rows, {
    widths: [9, 62, 16, 12, 22, 24],
    linkCols: [0], statusCols: [2], catCols: [4], initCols: [5], wrapCols: [1], rowHeight: 20,
  });
}

const out = require("path").join(OUT, "102326 Friday Email.xlsx");
wb.xlsx.writeFile(out).then(() => console.log("WROTE", out));
