# GreenFeed QC Dashboard (Next.js / Vercel)

This is a Vercel-friendly JavaScript conversion of your Streamlit GreenFeed QC app.

## What it does
- Authenticates to the C-LOCK API through a server route
- Downloads GreenFeed visits for one or more unit IDs
- Uploads and links `MVH Research.csv` in the browser
- Optionally uploads and links `Treatments.xlsx`
- Builds QC plots in the browser with Plotly
- Exports:
  - merged CSV
  - unit summary CSV
  - PDF report of the rendered charts

## Files
- `app/page.js` → main UI
- `app/api/greenfeed/route.js` → server-side proxy for C-LOCK login/data fetch
- `lib/greenfeed.js` → shared parsing, normalization, summaries

## Local run
```bash
npm install
npm run dev
```

## Deploy to Vercel
1. Upload this folder to GitHub.
2. Import the repo into Vercel.
3. Deploy.

No Vercel environment variables are required because username/password are entered in the UI and sent to the server route per request.

## Notes
This version reproduces the main workflow in JavaScript, but it is not a 1:1 replica of every Streamlit behavior. The charts and downloads are implemented client-side for Vercel compatibility.
