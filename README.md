# cp-register-template

Starter toolkit for creating a CP items assist register and pre-filling TRU information from available source data.

The repository now includes a Python generator that converts a TRU CSV file into a formatted Excel workbook with:

- `CP Register` sheet for CP item tracking
- `TRU Source Data` sheet preserving the imported source rows
- hidden `Lookups` sheet for controlled values
- formatted headers, filters, frozen panes, column sizing, and basic dropdown validation

## Repository contents

| Path | Purpose |
| --- | --- |
| `scripts/create_cp_register_template.py` | Generates the Excel CP register workbook from TRU CSV data. |
| `examples/tru_items.csv` | Small sample input file showing the expected CSV format. |
| `requirements.txt` | Python dependency list. |
| `.gitignore` | Ignores Python cache files, virtual environments, and generated workbooks. |

## Quick start

1. Create and activate a Python virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Generate a sample workbook:

```bash
python scripts/create_cp_register_template.py examples/tru_items.csv -o cp_register_template.xlsx
```

4. Open `cp_register_template.xlsx` in Excel and continue completing the CP register.

## Input CSV format

The generator accepts a UTF-8 CSV file. Required columns:

- `tru_tag`
- `tru_description`

Optional columns:

- `discipline`
- `system`
- `subsystem`
- `location`
- `priority` (`High`, `Medium`, or `Low`)
- `responsible_party`
- `source_reference`
- `comments`

If `priority` is blank, the script defaults it to `Medium`.

## Output workbook columns

The CP register sheet includes these columns:

- CP Item ID
- Discipline
- System
- Subsystem
- TRU Tag
- TRU Description
- CP Activity
- Location
- Priority
- Responsible Party
- Status
- Required Evidence
- Source Reference
- Comments

## Next improvements

Useful follow-up additions would be:

- company/project-specific CP activity rules
- import from Excel as well as CSV
- validation against an official TRU master list
- a Streamlit or web interface for non-technical users
- generated summary dashboard by priority, status, system, and responsible party
