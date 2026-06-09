# Claude Code — Project Instructions

## Git workflow
- After every commit, immediately push to `personal main` (`git push -u personal main`).
- Do not wait for the user to ask to publish.

## Remote
- Remote name: `personal`
- Branch: `main`

## Google Sheet data source
- Apps Script deployment URL: `https://script.google.com/macros/s/AKfycbzMPK8VXZ0nMSJWSP6_dMSabXrgksCNH4eP5RfopdARJXcMvrDriFQKOrckXwSuUZP4/exec`
- Sheet URL: `https://docs.google.com/spreadsheets/d/1rdQItcwoLnKBMzcpEpdWKk6eevwI8bQAqmICl_syoaw/edit`
- Write secret: `ey-pipeline-2026-ivan` (also in gitignored `data/.pipeline-secret`)
- Read:  `python3 data/pipeline-sheets.py read <Tab> <Range>`
- Write: `python3 data/pipeline-sheets.py setRow <Tab> '<match_json>' '<data_json>'`
- Available actions: `read` · `setRow` · `appendRow` · `deleteRow`
- Credentials: `data/.pipeline-credentials.json` (gitignored, service account)
- Examples:
  - `python3 data/pipeline-sheets.py read Engagements A1:G5`
  - `python3 data/pipeline-sheets.py setRow Engagements '{"account_key":"qiddiya","type":"A","saleMonth":"3"}' '{"prob":"0.55"}'`
  - `python3 data/pipeline-sheets.py appendRow Engagements '{"account_key":"expo","type":"CP","saleMonth":"5","duration":"4","grossK":"80","prob":"0.8","isY3Origination":"FALSE"}'`
