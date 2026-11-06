# Weekly (Monday/Friday) Email Process

Produces the weekly sprint recap email + styled Excel workbook from Azure DevOps
(`dev.azure.com/oneCELA`, project **TSI Sales and Ops**).

Example output: [`../week43/`](../week43/) (HTML + TXT email drafts, XLSX workbook).

## Prerequisites

- **Node.js** (any recent LTS)
- **Your own Azure DevOps access token** for `dev.azure.com/oneCELA`
  (Entra access token via `az account get-access-token --resource 499b84ac-1321-427f-aa17-267ca6975798`,
  or a personal PAT with *Work Items: Read*). **Never commit tokens to this repo.**
- Install dependencies once: `cd email-pipeline && npm install`

## Pipeline (run from `email-pipeline/`)

Set the token in your shell first (PowerShell):

```powershell
$env:ADO_TOK = (az account get-access-token --resource 499b84ac-1321-427f-aa17-267ca6975798 --query accessToken -o tsv)
```

1. **`node fetch.js`** — queries ADO for all User Stories in the target iteration
   (including items that were *ever* in it), walks revision history to flag
   *sprint-agreed* vs *mid-week additions*, and resolves the Feature → Epic →
   Initiative chain. Writes `wk43raw.json`.
2. **`node normalize.js`** — flattens the raw data into one record per story
   (state, owner, estimates, initiative, sprint-agreed flag, business impact)
   and prints quick stats. Writes `week43final.json`.
3. **`node build.js [outdir]`** — builds the styled Excel workbook
   (dashboard, status/category/lens color coding, clickable ADO links,
   capacity vs. completed hours). `style.js` is an earlier variant kept for reference.
4. **Email draft** — the HTML/TXT email is drafted by Copilot from
   `week43final.json`, using the previous week's email (see `week43/`) as the
   template: summary stats, completed work by initiative, spillovers with notes,
   and capacity adjustments for absences.

## Per-week edits

These scripts are a working snapshot from Week 43; each week you (or Copilot) update:

- `ITER` in `fetch.js` — iteration path, e.g. `TSI Sales and Ops\FY27 - Q2\Week 44 - 2026`.
- The Tuesday-cutoff timestamps (`TUE` in `fetch.js`, `CUT` in `normalize.js`) —
  stories entering the iteration after Tuesday 00:00 ET count as mid-week additions.
- In `build.js`: the `CATEGORY` / `RBT` (Run-Build-Transform) classification maps per
  story ID, `SPILL_NOTE` comments, and the capacity constants
  (`BASE_CAP`, `ABSENCES` — see `../user-stories/team-capacity-baseline.csv`).

In practice the easiest path is to open this repo in Copilot and ask it to
"run the Friday email for Week NN", pointing it at this document.
