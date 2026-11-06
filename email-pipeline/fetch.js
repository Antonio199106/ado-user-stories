// Fetch ADO Week 43 - 2026 (Oct 26-30) data
const https = require("https");
const fs = require("fs");
const TOK = process.env.ADO_TOK;
const ITER = "TSI Sales and Ops\\FY27 - Q2\\Week 43 - 2026";
const BASE = "dev.azure.com";

function req(method, path, body) {
  return new Promise((res, rej) => {
    const data = body ? JSON.stringify(body) : null;
    const r = https.request({
      host: BASE, path, method,
      headers: {
        Authorization: "Bearer " + TOK,
        "Content-Type": "application/json",
        ...(data ? { "Content-Length": Buffer.byteLength(data) } : {}),
      },
    }, resp => {
      let b = "";
      resp.on("data", c => b += c);
      resp.on("end", () => {
        if (resp.statusCode >= 300) return rej(new Error(resp.statusCode + " " + b.slice(0, 300)));
        res(JSON.parse(b));
      });
    });
    r.on("error", rej);
    if (data) r.write(data);
    r.end();
  });
}

const FIELDS = [
  "System.Id", "System.WorkItemType", "System.Title", "System.State", "System.AssignedTo",
  "System.Tags", "System.CreatedDate", "System.Parent", "System.IterationPath", "System.Description",
  "Microsoft.VSTS.Scheduling.OriginalEstimate", "Microsoft.VSTS.Scheduling.RemainingWork",
  "Microsoft.VSTS.Scheduling.CompletedWork", "Custom.DescribeBusinessImpact",
];

(async () => {
  // 1. current members of the iteration
  const wiql = await req("POST", "/oneCELA/TSI%20Sales%20and%20Ops/_apis/wit/wiql?api-version=7.0", {
    query: `SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = 'TSI Sales and Ops' AND [System.IterationPath] = '${ITER}' AND [System.WorkItemType] = 'User Story'`,
  });
  const ids = wiql.workItems.map(w => w.id);
  console.log("in iteration:", ids.length);

  // 2. try WAS EVER to catch items moved out during the week
  let everIds = [];
  try {
    const ever = await req("POST", "/oneCELA/TSI%20Sales%20and%20Ops/_apis/wit/wiql?api-version=7.0", {
      query: `SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = 'TSI Sales and Ops' AND [System.IterationPath] WAS EVER '${ITER}' AND [System.WorkItemType] = 'User Story'`,
    });
    everIds = ever.workItems.map(w => w.id);
    console.log("ever in iteration:", everIds.length);
  } catch (e) { console.log("WAS EVER unsupported:", e.message); }

  const all = [...new Set([...ids, ...everIds])];

  // 3. batch fetch fields
  const items = [];
  for (let i = 0; i < all.length; i += 200) {
    const b = await req("POST", "/oneCELA/_apis/wit/workitemsbatch?api-version=7.0", { ids: all.slice(i, i + 200), fields: FIELDS });
    items.push(...b.value);
  }

  // 4. for items created before Tue Oct 27 00:00 ET (2026-10-27T04:00Z), find when they entered the iteration
  const TUE = Date.parse("2026-10-27T04:00:00Z");
  for (const it of items) {
    const f = it.fields;
    const created = Date.parse(f["System.CreatedDate"]);
    it.enteredIter = null;
    if (created >= TUE) { it.sprintAgreed = false; continue; }
    // walk updates for IterationPath transitions into ITER
    const ups = await req("GET", `/oneCELA/TSI%20Sales%20and%20Ops/_apis/wit/workitems/${it.id}/updates?api-version=7.0&$top=200`);
    let entered = null;
    for (const u of ups.value) {
      const ch = u.fields && u.fields["System.IterationPath"];
      if (u.rev === 1) {
        const iv = ch ? ch.newValue : null;
        if (iv === ITER) entered = u.revisedDate || f["System.CreatedDate"];
      } else if (ch && ch.newValue === ITER) {
        entered = u.fields["System.ChangedDate"] ? u.fields["System.ChangedDate"].newValue : u.revisedDate;
      }
    }
    it.enteredIter = entered;
    it.sprintAgreed = entered ? Date.parse(entered) < TUE : created < TUE;
    process.stdout.write(".");
  }
  console.log("");

  // 5. parents (features) for initiative mapping
  const parents = [...new Set(items.map(i => i.fields["System.Parent"]).filter(Boolean))];
  const pItems = [];
  for (let i = 0; i < parents.length; i += 200) {
    const b = await req("POST", "/oneCELA/_apis/wit/workitemsbatch?api-version=7.0", { ids: parents.slice(i, i + 200), fields: ["System.Id", "System.Title", "System.WorkItemType", "System.Parent"] });
    pItems.push(...b.value);
  }
  // epics for any features whose parent unknown
  const epicIds = [...new Set(pItems.map(p => p.fields["System.Parent"]).filter(Boolean))];
  const eItems = [];
  for (let i = 0; i < epicIds.length; i += 200) {
    const b = await req("POST", "/oneCELA/_apis/wit/workitemsbatch?api-version=7.0", { ids: epicIds.slice(i, i + 200), fields: ["System.Id", "System.Title", "System.WorkItemType", "System.Parent"] });
    eItems.push(...b.value);
  }
  const initIds = [...new Set(eItems.map(p => p.fields["System.Parent"]).filter(Boolean))];
  const iItems = [];
  for (let i = 0; i < initIds.length; i += 200) {
    const b = await req("POST", "/oneCELA/_apis/wit/workitemsbatch?api-version=7.0", { ids: initIds.slice(i, i + 200), fields: ["System.Id", "System.Title", "System.WorkItemType"] });
    iItems.push(...b.value);
  }

  fs.writeFileSync("wk43raw.json", JSON.stringify({ inIteration: ids, items, features: pItems, epics: eItems, initiatives: iItems }, null, 1));
  console.log("items:", items.length, "features:", pItems.length, "epics:", eItems.length, "initiatives:", iItems.length);
  console.log("sprintAgreed:", items.filter(i => i.sprintAgreed).length, "midweek:", items.filter(i => !i.sprintAgreed).length);
})().catch(e => { console.error("FATAL", e.message); process.exit(1); });
