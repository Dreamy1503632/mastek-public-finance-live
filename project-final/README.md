# Mastek Public Finance Live 2026

## Structure
- `api/index.js` - Express app (all routes), exported for Vercel serverless
- `vercel.json` - routes all requests to api/index.js, keeps existing page URLs working
- `avatar.html`, `assessment.html`, `book-session.html` - front-end pages
- `.env.example` - template for required environment variables (copy to `.env` for local testing; set the same names in the Vercel dashboard for production)

## Not included here (retired / never committed)
- `server.js` - old local Express entry point (`.listen()`), replaced by `api/index.js`
- `azure.json`, `email.json`, `api.json` - old local credential files, replaced by environment variables

## Required environment variables
Set these in Vercel → Project → Settings → Environment Variables (see `.env.example` for the full list):
- AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_KEY, AZURE_OPENAI_DEPLOYMENT, AZURE_OPENAI_API_VERSION
- EMAIL_HOST, EMAIL_PORT, EMAIL_SECURE, EMAIL_USER, EMAIL_PASS, EMAIL_FROM, EMAIL_FROM_NAME
- BLOB_READ_WRITE_TOKEN is injected automatically once you connect a Blob store - no need to set manually

## Storage
Leads and bookings are appended as JSONL to Vercel Blob (`leads.jsonl`, `bookings.jsonl`).
Download a real spreadsheet any time by visiting:
- /api/export-leads
- /api/export-bookings

## Deploy
1. Push this repo to GitHub
2. Import into Vercel
3. Storage tab → Create Database → Blob → connect to project
4. Add all environment variables above
5. Redeploy
