# Vercel Excel Search (Serverless)

**What this is:** A Vercel-ready project that serves a static frontend and three serverless API routes that read an Excel file, preview rows, and download filtered results as a new `.xlsx` — all **free** on Vercel.

## Project structure
```
vercel-excel-search/
├─ api/
│  ├─ data.js      # GET /api/data?sheet=...
│  ├─ search.js    # GET /api/search?q=term&sheet=...
│  └─ sheets.js    # GET /api/sheets
├─ data/
│  └─ data.xlsx    # your Excel file (replace with your own)
├─ public/         # (optional) static assets
├─ index.html      # frontend UI
└─ package.json    # xlsx dependency
```

## Local preview (optional)
You can serve index.html with any static server, but the API routes need Vercel.
```bash
npm i -g vercel
vercel dev
# open http://localhost:3000
```

## Deploy to Vercel (free)
1) Push this folder to a GitHub repo.  
2) Go to https://vercel.com → "New Project" → "Import Git Repository".  
3) Pick your repo and click "Deploy".  
4) Open your Vercel URL → you're live!

## Using your own Excel
Replace `data/data.xlsx` with your file and commit.  
Or set `EXCEL_PATH` as a Vercel env var to point somewhere else.

## Notes
- Serverless functions read from the file packaged in your repo.
- For large files, consider CSV + pagination or a small DB.
