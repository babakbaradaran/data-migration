# Excel Job Mapping Splitter

This Python script automates the process of splitting an anonymized Excel mapping file containing multiple job entries into separate Excel files — one for each unique old_job.

Each output file contains only the relevant mapping details for that specific job, making it easier to process or track activity codes job-by-job.

---

## Input Requirements

The script scans the current folder for .xlsx files that:

- **Do not start with** m (to avoid reprocessing outputs)
- **Are not temp/backup files** (~$)
- Contain a worksheet named Sheet1
- Contain the following required columns:
  - old_job (Chose this because the column header in most of the mapping files was old_job just to know)
  - PHASE
  - Old Phase Code Description
  - New Phase
  - New Phase Description

---

## How It Works

1. Loops through all Excel files in the same folder as the script.
2. For each file:
   - Reads Sheet1 using pandas.
   - Verifies that all required columns are present.
   - Finds all unique old_job values.
   - Filters rows for each job.
   - Writes each job to its own file named like:
     
     mOLDJOB-xxxxx.xlsx
     
3. Each output file contains:
   - PHASE
   - Old Phase Code Description
   - New Phase`
   - New Phase Description

---

## Example

**Input file**: multi_job_anonymized_corrected.xlsx

| old_job       | PHASE           | Old Phase Code Description | New Phase        | New Phase Description |
|---------------|------------------|-----------------------------|------------------|------------------------|
| OLDJOB-00010  | P0000.00000001   | Phase Desc 1               | N0001.00000001   | New Phase Desc 1       |
| OLDJOB-00010  | P0000.00000002   | Phase Desc 2               | N0002.00000002   | New Phase Desc 2       |
| OLDJOB-00011  | P0000.00000001   | Phase Desc 1               | N0001.00000001   | New Phase Desc 1       |

**Output files**:
- mOLDJOB-00010.xlsx
- mOLDJOB-00011.xlsx

---

## Notes

- This version is based on anonymized data, safe for testing, training, and GitHub publishing.
- It is especially useful for separating mapping logic per job in systems like ERP/Project Controls/Cost Management.

---

## How to Use

Make sure Python and dependencies are installed:

```bash
pip install pandas openpyxl



