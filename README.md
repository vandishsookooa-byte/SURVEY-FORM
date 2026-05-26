# SURVEY-FORM

Professional, shareable Transport Department technical survey form for **RT Knits LTD**.

## How to use

1. Open `/home/runner/work/SURVEY-FORM/SURVEY-FORM/index.html` in any modern browser.
2. Fill in the two main sections:
   - Main Section A: Conveyance (for those concerned)
   - Main Section B: Production
3. Submit the form.
4. Each submission is saved in browser local storage and an updated Excel file is automatically generated:
   - `RT_Knits_LTD_Transport_Technical_Survey.xls`

## Notes

- The form includes detailed conveyance questions and an initial production block.
- Excel export is handled directly in-browser (no backend required and no setup needed).

## Should this be linked to Excel using Python?

- **Current setup:** No Python required. The browser directly generates and downloads an Excel-compatible `.xls` file after each submission.
- **Use Python only if you want a central shared file/database** (multi-user real-time storage). In that case:
  1. Keep this HTML as frontend
  2. Send submissions to a Python backend API (Flask/FastAPI/Django)
  3. Save to database and/or update a central Excel file on the server