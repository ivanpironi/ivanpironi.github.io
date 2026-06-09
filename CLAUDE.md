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
- Read:  `curl -L "DEPLOYMENT_URL"`
- Write: `curl -L -X POST "DEPLOYMENT_URL" -H "Content-Type: application/json" -d '{"token":"ey-pipeline-2026-ivan","action":"setRow","tab":"Engagements","match":{"account_key":"qiddiya","type":"A","saleMonth":"3"},"data":{"prob":0.55}}'`
- Available actions: `setRow` · `appendRow` · `deleteRow` · `setCell`
