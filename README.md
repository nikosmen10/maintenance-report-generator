# Maintenance Log Analyzer & PDF Report Generator

Automated Python tool that reads an Excel-based maintenance log and generates a professional multi-page PDF report — with charts, fault analysis, equipment rankings, and actionable recommendations.

Built for engineering and maintenance teams where data lives in spreadsheets but nobody has time to analyze it manually.

---

## What it produces

- **KPI summary** — total events, downtime hours, repair cost, pending jobs, critical event count
- **Monthly trend chart** — maintenance events vs repair cost over time
- **Equipment cost ranking** — top assets by total repair spend
- **Priority breakdown** — pie chart of Critical / High / Medium / Low events
- **MTBF per equipment** — mean time between failures calculated automatically
- **Recurring fault detection** — flags any fault appearing 3+ times on the same machine
- **Technician workload table** — jobs and hours per engineer
- **Auto-generated recommendations** — based on the actual data, not generic advice

---

## Example output

![Report preview](preview.png)

---

## Use cases

- **Food & beverage processing** — separators, homogenizers, pumps, heat exchangers
- **Manufacturing plants** — production line equipment, compressors, conveyors
- **Facilities management** — HVAC, utilities, building assets
- **Any team** that tracks maintenance in Excel and wants instant reporting

---

## Requirements

Your Excel file needs these columns:

| Column | Description |
|---|---|
| `Date` | Date of the maintenance event |
| `Equipment` | Equipment name or ID |
| `Fault Description` | What went wrong |
| `Priority` | Critical / High / Medium / Low |
| `Technician` | Who carried out the work |
| `Downtime (hrs)` | Hours the equipment was down |
| `Repair Cost (EUR)` | Cost of repair |
| `Status` | Completed / Pending / In Progress |

A sample file (`maintenance_log.xlsx`) is included to test with immediately.

---

## Setup

**1. Clone the repo**
```bash
git clone https://github.com/nikosmen10/maintenance-report-generator.git
cd maintenance-report-generator
```

**2. Install dependencies**
```bash
pip install pandas openpyxl reportlab matplotlib
```

**3. Run**
```bash
python maintenance_analyzer.py your_file.xlsx your_report.pdf
```

That's it. The PDF is generated in the same folder.

---

## Example

```bash
python maintenance_analyzer.py maintenance_log.xlsx maintenance_report.pdf
```

```
Report saved to: maintenance_report.pdf
```

---

## Adapting to your data

The column names in your Excel file must match the expected names listed above. If your file uses different column names (e.g. `Cost` instead of `Repair Cost (EUR)`), open `maintenance_analyzer.py` and update the column references at the top of the `load_and_analyze()` function.

Currency is set to EUR by default. To change it, search for `EUR` in the script and replace with your currency symbol.

---

## Tech stack

| Library | Purpose |
|---|---|
| pandas | Data loading and analysis |
| openpyxl | Excel file reading |
| reportlab | PDF generation |
| matplotlib | Chart rendering |

---

## Author

**Nikos** — Industrial IoT & Automation Engineer | Robotics MSc  
Field service background in industrial processing equipment (separators, homogenizers, heat exchangers).  
I build Python automation tools for engineering and manufacturing teams.

[Upwork Profile](https://www.upwork.com/freelancers/nikosmen10)
