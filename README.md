# SURVEY-FORM

Transport Department survey form for **RT Knits LTD** with:
- Two sections: **Conveyance** and **Production**
- Anonymous response note on the UI
- Flask backend storage to an auto-updating Excel file

## Run locally with Flask

1. Open a terminal in:
   - `/home/runner/work/SURVEY-FORM/SURVEY-FORM`
2. Install dependencies:
   - `pip install -r requirements.txt`
3. Start the app:
   - `python app.py`
   - If port 5000 is already in use, run `PORT=5001 python app.py`
4. Open:
   - `http://127.0.0.1:5000`
   - Or the alternate port you set, for example `http://127.0.0.1:5001`
5. Submit the form.

## Excel output

- Excel file is automatically created/updated on each submission:
  - `/home/runner/work/SURVEY-FORM/SURVEY-FORM/RT_Knits_Transport_Survey_Responses.xlsx`
- The form UI shows only **Submit Survey** (no download button on the form).

## Department list

- Existing departments plus:
  - IT Department
  - HR Department
  - Security
  - QA