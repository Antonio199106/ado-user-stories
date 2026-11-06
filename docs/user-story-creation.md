# User Story Creation (ADO)

User Stories are created in `dev.azure.com/oneCELA`, project **TSI Sales and Ops**
(board: *TSI Sales and Ops Team*), via Copilot using the Azure CLI
(`az boards work-item create` / REST API). Authenticate first with `az login`
and `az devops configure --defaults organization=https://dev.azure.com/oneCELA`.

## Prompt template

Paste this into Copilot, filling in the brackets:

```text
IMPORTANT US
Find the board called TSI Sales and Ops Team.
The user story will be assigned to [ASSIGNEE]
The sponsor will be [SPONSOR]
Original Estimate should be [N] hours
Input start date: [START] and end date: [END]
Determine the correct Parent (Feature) — [hint, e.g. "the ad-hocs one"]
The idea of the User Story is: "[one-paragraph description of the work]"
Write:
1) Request description
2) Describe business impact
3) Acceptance Criteria
ALL BASED ON THE USER STORY GUIDE PROVIDED
```

## Conventions

- Iteration path follows `TSI Sales and Ops\FY27 - Q<n>\Week <NN> - <year>`.
- Every story gets: assignee, sponsor, Original Estimate (hours), start/end dates,
  a parent Feature (which maps up to Epic → Initiative), a request description,
  a *Describe Business Impact* field, and acceptance criteria.
- Acceptance criteria should be verifiable (e.g. "process ran successfully and
  Fernando was informed of the outcome").
- Team capacity baseline for allocation checks: [`../user-stories/team-capacity-baseline.csv`](../user-stories/team-capacity-baseline.csv).

## Related recurring asks (all doable via Copilot + az CLI)

- Audit Not Started / Blocked / In Progress / On Hold stories with target dates before Friday.
- Compare each consultant's capacity vs. Original Estimate of their stories
  (current utilization) and rebalance over/under-allocated consultants.
- Create the week's own stories from calendar/meetings plus the usual weekly tasks.
