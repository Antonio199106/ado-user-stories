# ADO User Stories — Team Workflows

Everything needed to run our Azure DevOps workflows for project
**TSI Sales and Ops** (`dev.azure.com/oneCELA`):

| Workflow | Where |
|---|---|
| ADO analysis (querying stories via Azure CLI) | [`chat-transcript.md`](./chat-transcript.md) — methodology + summary below |
| User Story creation | [`docs/user-story-creation.md`](./docs/user-story-creation.md) |
| Monday/Friday email + weekly workbook | [`docs/friday-email-process.md`](./docs/friday-email-process.md), scripts in [`email-pipeline/`](./email-pipeline/), example output in [`week43/`](./week43/) |
| Team capacity baseline | [`user-stories/team-capacity-baseline.csv`](./user-stories/team-capacity-baseline.csv) |

## Getting started (new team member)

1. Get access to the ADO project **TSI Sales and Ops** on `dev.azure.com/oneCELA`.
2. Install Azure CLI, run `az login`, then
   `az devops configure --defaults organization=https://dev.azure.com/oneCELA`.
3. Install Node.js (for the email pipeline) and run `npm install` inside `email-pipeline/`.
4. Use **your own** credentials/token everywhere (`ADO_TOK` env var for the pipeline).
   Never commit tokens, PATs, or raw data exports to this repo.

## Summary

Total User Stories found: **3,845**

| State | Count |
|---|---|
| Closed | 3,677 |
| Removed | 90 |
| Not Started | 32 |
| In Progress | 29 |
| Resolved | 6 |
| Blocked | 5 |
| Assessment | 4 |
| On Hold | 2 |
