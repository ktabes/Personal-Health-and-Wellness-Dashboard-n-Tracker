# Personal Health & Wellness Dashboard n Tracker (Google Sheets + Apps Script)

Centralized tracker for daily nutrition, hydration, supplements, medications, weight, and body fat percent. Built in Google Sheets with Apps Script automation. Designed to show practical analytics, clean data workflows, and documentation.

> **Live sheet:** https://docs.google.com/spreadsheets/d/1nwOKdNyirYvzneHgH5YshVqYjr6sZe4YDX70giHiwi0/edit?usp=sharing (view only)  
> **Code:** this repository  
> **Contact:** ktabesbusiness@gmail.com or https://www.linkedin.com/in/kyle-table/

---

## Contents

- [Features](#features)  
- [Sheet Structure](#sheet-structure)  
- [Tech](#tech)  
- [Quick Start](#quick-start)  
- [Automation and Triggers](#automation-and-triggers)  
- [KPIs](#kpis)  
- [Workflows](#workflows)  
- [Troubleshooting](#troubleshooting)  
- [Repo Layout](#repo-layout)
- [Roadmap](#roadmap)    
- [Acknowledgments](#acknowledgments)

---

## Features

- Structured daily logging for meals, drinks, supplements, medications, weight, and body fat percent  
- Automatic per entry calculations and daily rollups for calories, macros, and key nutrients  
- Autocomplete from reference tables, input validation, conditional formatting for data quality  
- Compliance tracking for supplements and medications with rolling rates  
- Time series for weight and body fat percent with rolling averages and entry over entry change  

---

## Sheet Structure

- **Health & Wellness Summary:** Overview KPIs, quick filters, last refresh, weekly charts for macros, hydration, supplement and medication compliance, weight and body fat percent trends.
- **Inputs:** Daily entry sheet for meals and drinks (with units), hydration, supplements taken, medications taken, weight, body fat percent, and notes. Uses dropdowns, units, and validation.
- **Data Tables:** Aggregated data from **Inputs**. Includes daily total nutrition and macros, daily supplements and medications summaries, daily hydration totals, and other day-level rollups used for analysis and pivots.
- **Time Tables:** Event timing only. Stores the time a meal was eaten, the time skincare was used, and other timestamps for time-of-day analysis.
- **Cardio Data:** Daily steps and distance only. Tracks steps and distance in miles and kilometers, with daily and monthly totals.
- **Nutrition Reference:** Master list of foods and drinks with macros and selected micronutrients. Powers autocomplete and unit conversions for **Inputs**.
- **Supplements:** Catalog of supplements with dosage, daily schedule, and related details. 
- **Skincare:** Products in use, when they are used (AM/PM or schedule), reason for use, and product links.
- **Change Log:** Date, change, reason, and impact. Documents schema or logic updates for transparency.

> You may also have hidden helper tabs for lookups or list rebuilding.

---

## Tech

- Google Sheets, Apps Script V8, AI (ChatGPT, Gemeni, Claude), Named ranges, Data validation, Conditional formatting  
- Triggers: onEdit, Custom Menu, optional time driven  
- Version control: GitHub for source, optional clasp for sync

---

## Quick Start

1. Open the live sheet, then select File, Make a copy  
2. Populate **Nutrition Reference** with items you use. Add macros and key micronutrients as available  
3. Populate **Supplements** with supplement names, dosages, and schedule. Populate **Skincare** with products, usage schedule, reasons, and links  
4. Begin logging in **Inputs** for meals, drinks, hydration, supplements taken, medications taken, weight, and body fat percent  
5. **Data Tables** will aggregate daily totals from Inputs. **Time Tables** will record event times when entries are logged  
6. For **Cardio Data**, setup GoogleFit API for automated tracking of daily steps and distance (in mi and/or km)  
7. Review **Health & Wellness Summary** for KPIs and charts

---

## Automation and Triggers

- **Custom Menu (primary driver):** A top-level menu controls all writes and refreshes. Each action validates inputs, writes to **Inputs**, updates **Data Tables**, and, when relevant, stamps **Time Tables**.
  - Submit All
  - Submit Weight
  - Submit Body Fat
  - Submit Nutrition
  - Submit Water
  - Submit Supplements
  - Submit Skincare
  - Submit Vyvnase
  
- **Inputs:** All entries are added via the Custom Menu. Each action:
  - Validates units and ranges, parses numbers safely.
  - Writes a normalized row to **Inputs** and stamps `Date` and `Time`.
  - Calls the aggregation routine to update **Data Tables**.
  - Writes event timestamps to **Time Tables** for time-of-day analysis where applicable (for example, meal time, skincare time).

- **Data Tables:** Aggregations are rebuilt and are called automatically after each input menu action. Rollups include:
  - Daily total nutrition and macros.
  - Daily hydration totals.
  - Daily supplement and medication compliance by item and overall. 

- **Time Tables:** Populated only by the menu actions that capture timing. Stores event times such as when a meal was eaten and when skincare was used. Uses an onEdit trigger to keep entries sorted descending by date.

- **Cardio Data:** Sourced from **Google Fit API** on a time-driven trigger (for example, hourly or daily).
  - The trigger fetches steps and distance, dedupes by date, and writes daily values.
  - Converts distance between miles and kilometers. 

- **Nutrition Reference:** Two ways to add items:
  1) Edit the **Nutrition Reference** sheet directly.  
  2) From **Inputs**, fill the Nutrition fields with a new item name and complete macro data, then add the item. If the item does not exist, the script appends it to **Nutrition Reference**.  
  A sheet specific **onEdit** keeps the list alphabetized and automatically maintains the named ranges used for autocomplete. 

-**Supplements and Skincare:** Reference sheets for your own tracking. They are not used to compute compliance directly.  
  - **Supplements:** when you log a supplement via the Custom Menu or Inputs, the script looks up the default dosage from the **Supplements** sheet and stamps that value into **Time Tables** along with the timestamp. Editing the **Supplements** sheet updates the default dosage used on future logs.  
  - **Skincare:** reference for products, when you use them, reasons, and links. Automations read timestamps from **Inputs** to **Time Tables** when you log a skincare event. The **Skincare** sheet itself is not used for compliance or calculations.

- **Change Log:** Manual notes for structure changes, formula updates, and target adjustments.

---

## KPIs

**Nutrition and hydration**
- Calories per day, protein, carbs, fat, fiber, sugar, sodium, potassium  
- Macro adherence percent versus targets  
- Fiber grams per 1000 kcal  
- Added sugar grams per day  
- Sodium milligrams per day and banding  
- Hydration liters per day and by week

**Supplements and medications**
- Compliance percent by item and overall  
- Missed doses by day and by item  
- Average time of dose versus planned schedule

**Weight and body fat percent**
- 7 day rolling averages  
- Week over week change  
- Simple goal deltas to target weight or body fat percent

**Quality flags**
- Invalid unit or out of range values  
- Duplicates or missing required fields  
- Reconciliation checks that per entry sums match daily totals

---

## Workflows

- **Add or update Nutrition Reference:** add items either on the **Nutrition Reference** sheet or by entering a new item in **Inputs** with its macro data. A sheet-specific onEdit keeps the list alphabetized and maintains the named ranges used for autocomplete. No manual dropdown refresh is needed.
- **Daily logging via Custom Menu:** use the menu actions (Add Meal, Add Drink, Log Supplement, Log Medication, Log Weight or Body Fat %) to write normalized rows to **Inputs**. Each action validates units and ranges, updates **Data Tables**, and writes event timestamps to **Time Tables** where applicable.
- **Cardio Data:** fetched automatically from the Google Fit API on a time-driven trigger. The job dedupes by date, records steps and distance, and converts between miles and kilometers. No manual entry required. 
- **Supplements and Skincare catalogs:** maintain items, default dosages, schedules, reasons, and links. The **Supplements** sheet provides default dosage lookups when you log a supplement; **Skincare** is a reference list. Compliance and usage are calculated from **Inputs** and **Time Tables**, not directly from these catalogs.
- **Summary and analysis:** **Health & Wellness Summary** is a high-level snapshot that aggregates KPIs and charts for quick review. Dedicated analysis sheets will be added for deeper pivots, segmenting, and time-of-day views.
- **Weekly review:** filter **Health & Wellness Summary** to the last 7 or 28 days, scan KPIs and trends, and record any target or structure changes in the **Change Log**.


---

## Troubleshooting

- **Custom Menu is missing**
  - Close and reopen the sheet to trigger `onOpen`.  
  - In Apps Script, run the menu builder function once to authorize.  
  - Check that the script project is bound to this spreadsheet.

- **Menu action runs but nothing appears in Inputs**
  - Confirm required fields are filled and pass validation.  
  - Check that the menu function writes to the correct tab name `Inputs`.  
  - Look for protected ranges or filter views blocking writes.

- **Time Tables is empty**
  - Only menu actions that capture timing will write here. Log an event via the menu.  
  - Verify the write function targets the correct Time Tables columns.  
  - Ensure that you included a time in your input, not just the date.

- **Nutrition Reference did not pick up a new item**
  - Adding via Inputs: ensure you provided a new item name, selected the correct category (Food or Drink), and full macro data before submitting.  
  - Adding directly on Nutrition Reference: the sheet-level onEdit must be enabled and not blocked by protection.  
  - If items look unsorted or lists stale, confirm the onEdit sort is active and named ranges cover the new rows.

- **Dropdowns or autocomplete do not show new items**
  - Confirm the source named ranges point at Nutrition Reference and include the new rows.  
  - Avoid editing inside a filtered view when adding items.  
  - Reopen the sheet if lists appear cached.

- **Cardio Data did not refresh**
  - Time-driven trigger: check it exists and is enabled in Triggers.  
  - Google Fit API: confirm authorization and scopes, and that the correct Google account is used.  
  - Dedupe key: verify the script does not skip today’s entry due to a duplicated date.  
  - Unit conversions: check miles and kilometers formulas and input units.

- **Supplements or Skincare catalogs not reflected in logging**
  - For Supplements, default dosage is read when you log via the menu. Make sure the item name matches exactly.  
  - Skincare is a reference list. Usage timestamps are captured from Inputs, not from the catalog sheet.

- **KPIs look wrong**
  - Check for non-numeric values or unit text in numeric columns.  
  - Inspect red flagged cells for validation errors.  
  - Verify daily totals equal the sum of entries for that date.

- **General tips**
  - Remove merged cells in any write target range.  
  - Verify exact tab names match what the script expects.  
  - Check locale settings if decimal commas are used; functions may expect dots.

---

## Repo Layout

/src
appsscript.json
Cardio Tracker.gs
Inputs.gs
README.md

---

## Roadmap

- **Analysis suite inside the workbook**  
  Goal: build one or two dedicated analysis tabs that tell a clear story without leaving Sheets.  
  Approach: create normalized export tables in Data Tables, then layer pivots and charts with slicers.  
  Visuals and views:
  - Nutrition and adherence: weekly macro adherence, fiber per 1000 kcal, hydration by day, supplement and medication compliance by item and overall
  - Timing and outcomes: meal time distribution from Time Tables, late day calories vs next day weight delta, supplement time vs adherence
  - Pareto of top foods driving calories, box plot of added sugar by weekday, rolling 7 and 28 day trends for weight and body fat percent
  - Slicers for date window, weekday, meal type, supplement item  
  Deliverables: two Analysis sheets, saved pivot configurations, consistent theme and legend, short insights box on each sheet.  
  Success criteria: a reviewer can change the date window and see all visuals update in under 3 seconds.

- **Live Tableau and Power BI dashboards fed by the sheet**  
  Goal: keep external BI dashboards automatically up to date from your Google Sheet data.  
  Data model to expose:  
  - fact_nutrition_daily, fact_hydration_daily, fact_supplement_events, fact_medication_events, fact_cardio_daily  
  - dim_date, dim_item, dim_supplement, dim_medication  
  Tableau path:
  1) Make a read only “Exports” tab per fact table with headers in row 1 only.  
  2) Use Tableau Cloud Desktop or Web to connect to Google Sheets and select the Exports tabs.  
  3) Publish to Tableau Cloud, set a refresh schedule that matches your time driven jobs.  
  Power BI path:
  1) Publish each Exports tab to the web as CSV, or use an Apps Script that writes stable CSV files to Drive and exposes direct download links.  
  2) In Power BI Desktop use Get Data, Web, paste each CSV URL, set column types, publish to Power BI Service.  
  3) Configure scheduled refresh and credentials. If using public CSV, no gateway is required.  
  Deliverables: one Tableau dashboard and one Power BI dashboard with matching visuals and a README section listing the CSV endpoints and table contracts.  
  Success criteria: BI dashboards reflect a new Inputs entry within the next scheduled refresh without manual steps.

- **Automation and reliability pack**  
  Goal: push more work off your hands and make failures obvious and recoverable.  
  Additions:
  - Scheduler: time driven triggers for Google Fit fetch, nightly rebuild of Data Tables, weekly archive of older logs if used
  - Idempotent writes: stable keys like {date, item, source} to avoid duplicates in Data Tables and Cardio Data
  - Retry and backoff: simple try and catch with two retries for API calls, status to a Logs sheet with timestamp, action name, rows affected, duration, success or error
  - Health checks: red flag cells for missing required fields, non numeric values, and reconciliation checks where day totals must equal the sum of entries
  - One click menu actions: Rebuild Data Tables, Recalculate KPIs, Fetch Google Fit Cardio now, Backfill Cardio for a date range  
  Deliverables: Logs sheet with last 30 runs, menu actions wired, trigger list documented in the README, error messages that surface in toasts.  
  Success criteria: zero manual edits needed for daily operation, clear error context when an API or write fails, and average rebuild under 5 seconds on your current dataset.

---

## Acknowledgments

Built with **Google Sheets** + **Apps Script** + **AI (ChatGPT, Gemeni, Claude)**.
