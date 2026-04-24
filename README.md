# GreenFeed QC Dashboard (Next.js / Vercel)

This is a Vercel-friendly JavaScript version of your Streamlit GreenFeed QC app.

## What it does
- Authenticates to the C-LOCK API through a server route
- Downloads GreenFeed visits for one or more unit IDs
- Uploads and links `MVH Research.csv` in the browser
- Optionally uploads and links `Treatments.xlsx`
- Flags unmatched GreenFeed RFID rows that fail MVH linkage
- Filters matched data by unit, treatment, day range, and animal search
- Builds QC plots in the browser with Plotly
- Exports:
  - filtered merged CSV
  - unit summary CSV
  - daily summary CSV
  - unmatched rows CSV
  - PDF report of the selected charts

## Files
- `app/page.js` -> main UI, filtering, previews, and exports
- `app/api/greenfeed/route.js` -> server-side proxy for C-LOCK login and data fetch
- `lib/greenfeed.js` -> shared parsing, normalization, and summaries

## Local run
```bash
npm install
npm run dev
```

## Deploy to Vercel
1. Upload this folder to GitHub.
2. Import the repo into Vercel.
3. Deploy.

No Vercel environment variables are required because username and password are entered in the UI and sent to the server route per request.

## Notes
This version reproduces the main workflow in JavaScript, but it is not a 1:1 replica of every Streamlit behavior. Charts and downloads are implemented client-side for Vercel compatibility.

Browser preferences such as unit IDs, date range, filenames, and whether treatment uploads are enabled are saved locally. Passwords are not stored.
