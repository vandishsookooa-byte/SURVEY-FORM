# SURVEY-FORM

Generate a professional management slide deck from survey responses focused on Conveyance and Production.

## Features

The script builds a 6-slide presentation:

1. Executive Summary
   - Total responses
   - Response rate (when headcount is available)
   - Overall conveyance and production averages
   - Favorable percentages (ratings 4-5)
2. Conveyance Performance
   - Average score by metric
   - Rating distribution (1-5)
   - Department-wise conveyance heat table
3. Production Performance
   - Average score by metric
   - Rating distribution (1-5)
   - Department-wise production heat table
4. Conveyance vs Production Comparison
   - Side-by-side average comparison
   - Favorable % comparison
   - Department delta (Production - Conveyance)
5. Voice of Employee
   - Top positive themes
   - Top pain points
   - 3-5 anonymized comment quotes
6. Actions & Ownership
   - Top 5 improvement actions
   - Owner, priority, target, and KPI

## Expected Input Columns

The generator supports these survey columns (case/spacing-insensitive matching):

- `employeeName`
- `employeeCode`
- `departments`
- `conveyanceAllocationRating`
- `specialTransportRating`
- `conveyanceResponsivenessRating`
- `overallConveyanceSatisfaction`
- `productionPlanningRating`
- `logisticsRequestRating`
- `productionResponsivenessRating`
- `overallProductionSatisfaction`
- `comments`

It automatically:
- Cleans blanks and `N/A`
- Standardizes department names
- Converts rating fields to numeric 1-5

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python generate_management_slides.py \
  --input /absolute/path/to/survey_responses.xlsx \
  --output /absolute/path/to/Survey_Management_Deck.pptx \
  --company "RT Knits"
```

Supported input formats: `.xlsx`, `.xls`, `.xlsm`, `.csv`.
