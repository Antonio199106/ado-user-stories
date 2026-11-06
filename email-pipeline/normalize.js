// Normalize wk43raw.json -> week43final.json with initiative chain + stats scaffold
const fs = require("fs");
const R = JSON.parse(fs.readFileSync("wk43raw.json"));
const CUT = Date.parse("2026-10-28T04:00:00Z"); // end of Tue Oct 27 ET

const featById = {}, epicById = {}, initById = {};
R.features.forEach(f => featById[f.id] = f);
R.epics.forEach(e => epicById[e.id] = e);
R.initiatives.forEach(i => initById[i.id] = i);

const strip = h => (h || "").replace(/<style[\s\S]*?<\/style>/gi, "").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&gt;/g, ">").replace(/&lt;/g, "<").replace(/&#\d+;/g, " ").replace(/\s+/g, " ").trim();

const SHORT = {
  "Reporting and Automation Support": "Reporting & Automation",
  "Sales Excellence Analytics and Operations": "Sales Excellence",
  "In-Year Parenting / Regional Sales Ops (Americas)": "In-Year Parenting",
  "Employee Giving": "Employee Giving",
};

const out = R.items.map(it => {
  const f = it.fields;
  const pid = f["System.Parent"];
  const feat = pid ? featById[pid] : null;
  const epic = feat && feat.fields["System.Parent"] ? epicById[feat.fields["System.Parent"]] : null;
  const init = epic && epic.fields["System.Parent"] ? initById[epic.fields["System.Parent"]] : null;
  const initName = init ? init.fields["System.Title"] : (epic ? epic.fields["System.Title"] : "Unmapped");
  const created = f["System.CreatedDate"];
  const sa = Date.parse(created) < CUT && (!it.enteredIter || Date.parse(it.enteredIter) < CUT);
  return {
    id: it.id,
    title: f["System.Title"],
    state: f["System.State"],
    assigned: f["System.AssignedTo"] ? f["System.AssignedTo"].displayName : "",
    tags: f["System.Tags"] || "",
    created: created.slice(0, 10),
    est: f["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0,
    rem: f["Microsoft.VSTS.Scheduling.RemainingWork"] || 0,
    parent: pid || "",
    feature: feat ? feat.fields["System.Title"] : "",
    epic: epic ? epic.fields["System.Title"] : "",
    initiative: SHORT[initName] || initName,
    sprintAgreed: sa,
    inIteration: R.inIteration.includes(it.id),
    impact: strip(f["Custom.DescribeBusinessImpact"]).slice(0, 500),
    desc: strip(f["System.Description"]).slice(0, 300),
  };
});
fs.writeFileSync("week43final.json", JSON.stringify(out, null, 1));

// quick stats
const sa = out.filter(o => o.sprintAgreed), mw = out.filter(o => !o.sprintAgreed);
const closed = out.filter(o => ["Closed", "Resolved"].includes(o.state));
const sum = a => Math.round(a.reduce((s, o) => s + o.est, 0) * 10) / 10;
console.log("SA:", sa.length, sum(sa), "hrs | closed", sa.filter(o => ["Closed", "Resolved"].includes(o.state)).length);
console.log("MW:", mw.length, sum(mw), "hrs | closed", mw.filter(o => ["Closed", "Resolved"].includes(o.state)).length);
console.log("Closed total:", closed.length, sum(closed), "hrs");
console.log("Open:", out.filter(o => !["Closed", "Resolved"].includes(o.state)).map(o => `${o.id} ${o.state} ${o.est}h rem${o.rem} ${o.title.slice(0, 50)}`));
console.log("Not in iteration (moved out):", out.filter(o => !o.inIteration).map(o => o.id));
const byInit = {};
closed.forEach(o => { byInit[o.initiative] = (byInit[o.initiative] || 0) + o.est; });
console.log("Closed hrs by initiative:", byInit);
console.log("Unmapped:", out.filter(o => o.initiative === "Unmapped").map(o => o.id + " " + o.title.slice(0, 40) + " feat:" + o.feature));
