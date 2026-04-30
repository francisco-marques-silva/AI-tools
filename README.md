<p align="center">
  <img src="scripts/logo.avif" alt="AI-Tools Logo" width="200">
</p>

<h1 align="center">AI-Tools — Systematic Review Screening &amp; Validation Platform</h1>

> Automates title/abstract screening with OpenAI models and generates a comprehensive multi-project validation report — all from the browser. No command line required for the standard workflow.

---

## Table of Contents

1. [Overview](#overview)
2. [Quick Start](#quick-start)
3. [Web Interface Guide](#web-interface-guide)
   - [Screening Tab](#screening-tab)
   - [Report Tab — How to Use Validation](#report-tab--how-to-use-validation)
4. [File Naming Conventions](#file-naming-conventions)
5. [Report Sections (16 + Appendices)](#report-sections-16--appendices)
6. [Charts (16 publication-ready)](#charts-16-publication-ready)
7. [Advanced: CLI Usage](#advanced-cli-usage)
8. [Kappa Interpretation](#kappa-interpretation-landis--koch-1977)
9. [Troubleshooting](#troubleshooting)

---

## Overview

AI-Tools has two main features, both accessible from the browser:

| Feature | Where | What it does |
|---------|-------|--------------|
| **AI Screening** | Screening tab | Uploads a TIAB spreadsheet, sends every article to an OpenAI model, returns `include` / `exclude` / `maybe` decisions |
| **Validation Report** | Report tab | Compares AI decisions against human reviewers, computes sensitivity/specificity/kappa, generates Word report + 16 charts viewable in the browser |

---

## Quick Start

### Step 1 — Install Python

Download and install Python 3.12+ from [python.org](https://www.python.org/downloads/).

⚠️ During installation, check **"Add Python to PATH"** on the first screen.

### Step 2 — Open a Terminal

Open **PowerShell** (Windows) or **Terminal** (macOS/Linux) and navigate to the project folder:

```powershell
cd "C:\path\to\AI-tools"
```

### Step 3 — Create a Virtual Environment

```powershell
python -m venv .venv
```

### Step 4 — Activate the Virtual Environment

```powershell
# Windows PowerShell
.\.venv\Scripts\Activate.ps1

# Linux / macOS
source .venv/bin/activate
```

### Step 5 — Install Dependencies

```powershell
.\.venv\Scripts\pip.exe install -r requirements.txt
```

### Step 6 — Start the Web App

```powershell
.\.venv\Scripts\python.exe -m uvicorn scripts.backend:app --reload --port 8000
```

Open **http://localhost:8000** in your browser. Everything else is done through the interface.

### Every Time You Return

```powershell
cd "C:\path\to\AI-tools"
.\.venv\Scripts\Activate.ps1
.\.venv\Scripts\python.exe -m uvicorn scripts.backend:app --reload --port 8000
```

---

## Web Interface Guide

### Screening Tab

The **Screening** tab runs AI-assisted title/abstract screening.

1. **Configuration** — select a model (GPT-5, GPT-4.1, GPT-4o families) and enter your OpenAI API key.
2. **Study Context** — write your PICO synopsis and add inclusion/exclusion criteria (one per line, press Enter to add).
3. **Spreadsheet** — drag and drop your TIAB file (`.xlsx`, `.xls`, or `.csv`). It must have `title` and `abstract` columns. A preview of the first 5 rows is shown.
4. **Send to Backend** — starts the job. A real-time progress bar shows how many articles have been processed and current concurrency.
5. **Download** — when complete, download the result as **CSV** or **XLSX**. The file contains `screening_decision` (`include`/`exclude`/`maybe`) and `screening_reason` for each article.

Settings (model, API key, PICO, criteria) are saved automatically in the browser.

---

### Report Tab — How to Use Validation

The **Report** tab compares AI screening results against human reviewer decisions and generates a complete statistical validation report.

#### Step 1 — Run AI Screening

Use the **Screening** tab to screen your TIAB spreadsheet. Download the result.

#### Step 2 — Rename the AI result file

Rename the downloaded file to follow the convention:

```
YYYYMMDD - model - 1º teste - project.xlsx
```

| Part | Example | Description |
|------|---------|-------------|
| `YYYYMMDD` | `20260227` | Date the screening was run |
| `model` | `gpt-5-mini` | Model used (must match exactly) |
| `1º teste` | `1º teste` | Test number — use `2º teste` for a replication run |
| `project` | `zebra` | Your project name (used in all related files) |

#### Step 3 — Prepare human reference spreadsheets

Create the following files for each project (each must have at least `title` and `decision` columns):

| File | Pattern | Purpose |
|------|---------|---------|
| **TIAB** | `project - TIAB.xlsx` | Human decisions at title/abstract screening phase |
| **Fulltext** | `project - Fulltext.xlsx` | Articles selected for full-text reading |
| **Listfinal** | `project - Listfinal.xlsx` | Final included articles after full-text reading (gold standard) |

The `decision` column must contain `include`, `exclude`, or `maybe`.

> **Critical:** The `title` column in your human TIAB file must match the `title` column in your AI result file exactly (character for character). The easiest way to guarantee this is to use the same source spreadsheet you uploaded for AI screening.

#### Step 4 — Create metadata.xlsx

Create a file named `metadata.xlsx` with one row per AI execution:

| Column | Example | Description |
|--------|---------|-------------|
| `project` | `mino` | Project name — must match file naming |
| `code` | `20260227` | YYYYMMDD from the AI result filename |
| `model` | `gpt-5-mini` | Model name — must match AI result filename |
| `parameter` | `reasoning=medium` | Free-text description of parameters |
| `version` | `1º teste` | `1º teste` or `2º teste` |
| `time_ia` | `0:04:32` | AI execution time (H:MM:SS) |
| `time_human` | `2:00:00` | Human TIAB time: 1 min/record × 2 reviewers |
| `tokens input` | `152340` | Input tokens (from OpenAI dashboard) |
| `tokens_output` | `28450` | Output tokens |
| `cost_input` | `0.023` | Input cost (USD) |
| `cost_output` | `0.114` | Output cost (USD) |
| `cost_total` | `0.137` | Total cost (USD) |

You can include multiple projects and multiple model runs in the same file — one row per execution.

#### Step 5 — Upload and generate

1. Click the **Report** tab.
2. Drag all your input files into the **Input Files** area:
   - AI result file(s) — `YYYYMMDD - model - Xº teste - project.xlsx`
   - Human reference files — `project - TIAB.xlsx`, `project - Fulltext.xlsx`, `project - Listfinal.xlsx`
   - `metadata.xlsx`
3. Click **Generate Report**.

The system will:
- Run the full analysis pipeline
- Generate a **Word report** (16 sections, tables and text)
- Generate a **chart data XLSX** (14 sheets of source data)
- Generate **16 publication-ready charts**, viewable directly in the browser

When complete, the results appear below the log:
- **Downloads** — Word report (`.docx`) and chart data (`.xlsx`)
- **Charts** — all 16 charts in a clickable gallery (click any chart to enlarge)
- **Data Tables** — browse all 14 data sheets directly in the browser

#### Interpreting key metrics

| Metric | Target | Description |
|--------|--------|-------------|
| **Sensitivity** | ≥ 95% | % of human-included articles the AI also included |
| **Specificity** | ≥ 50% | % of human-excluded articles the AI also excluded |
| **Listfinal Capture** | ≥ 95% | % of final gold-standard articles the AI retained |
| **Test-Retest Kappa** | ≥ 0.80 | Reproducibility of AI decisions across runs |
| **F1 Score** | Higher is better | Harmonic mean of sensitivity and PPV |

---

## File Naming Conventions

### AI Result Files

```
YYYYMMDD - model - Xº teste - project.xlsx
```

**Required columns:** `title`, `screening_decision`

### Human Reference Files

```
project - TIAB.xlsx
project - Fulltext.xlsx
project - Listfinal.xlsx
```

**Required columns:** `title`, `decision` (plus `abstract` when available in TIAB)

### Multiple Projects

You can place files for multiple projects in a single upload. Just use consistent project names:

```
20260227 - gpt-5-mini - 1º teste - mino.xlsx
20260227 - gpt-5-mini - 2º teste - mino.xlsx
20260228 - gpt-4o - 1º teste - mino.xlsx
mino - TIAB.xlsx
mino - Fulltext.xlsx
mino - Listfinal.xlsx
20260301 - gpt-5-mini - 1º teste - zebra.xlsx
zebra - TIAB.xlsx
zebra - Fulltext.xlsx
zebra - Listfinal.xlsx
metadata.xlsx       ← contains rows for ALL projects
```

---

## Report Sections (16 + Appendices)

The report begins with a **Methodological Notes & Report Guide** section, then contains:

| Section | Content |
|---------|---------|
| 1. Data Validation | File inventory, metadata correspondence checks |
| 2. Metadata and Costs | Execution details, token counts, costs per model/project |
| 3. Diagnostic Analysis | Sensitivity, specificity, PPV, NPV, F1, Kappa vs human TIAB |
| 4. Fulltext Verification | Capture rate of articles sent to full-text reading |
| 5. Listfinal Verification | Capture rate of final included articles (gold standard) |
| 6. Test-Retest | Reproducibility between two independent runs (Kappa + CI) |
| 7. False Negatives | Articles included by humans but excluded by AI |
| 8. False Positives | Articles excluded by humans but included by AI |
| 9. General Comparative Table | All metrics consolidated in one view |
| 10. Cost-Effectiveness | Cost (USD) vs mean sensitivity per model |
| 11. Workload Reduction | Time saved (AI vs human), speed factor |
| 12. Absolute Efficiency | Selectivity × capture score |
| 13. Dual Gold Standard | Metrics using both TIAB and Listfinal as gold standards |
| 14. Combined Runs | Union sensitivity/specificity when merging two test runs |
| 15. Aggregated Performance | Mean ± SD across all projects and runs |
| 16. F1 vs Cost per Article | Scatter with regression line |
| Appendix A | False Positives TIAB — title, abstract, which models flagged |
| Appendix B | Missed Fulltext Articles — per-model matrix + abstract details |

---

## Charts (16 publication-ready)

All 16 charts are generated automatically and viewable in the browser. Click any chart to enlarge.

| Chart | Description |
|-------|-------------|
| Sensitivity by Model per Project | Grouped bars, individual run dots, 95%/80% reference lines |
| Listfinal Capture Rate Heatmap | Red-yellow-green heatmap across models × projects |
| Test-Retest Kappa | Bar chart with 95% CI, threshold reference lines |
| Model Comparison Radar | Spider chart across 6 key dimensions |
| Cost vs Sensitivity | Bubble chart (bubble size = F1 score) |
| Workload Reduction | Human vs AI time + speed factor |
| Efficiency Frontier (Grid) | Per-run scatter: AI positive rate vs Listfinal capture |
| Efficiency Frontier (by Project) | Per-project average efficiency frontier |
| Efficiency Frontier (Overall) | All-project aggregated efficiency frontier |
| Efficiency Score by Project | Bar chart per model per project |
| Efficiency Score Aggregated | Mean ± SD efficiency score |
| Sensitivity & Specificity (Dual Gold) | Paired bars — TIAB vs Listfinal gold standards |
| Aggregated Performance | 2×3 subplot grid: bar + dot overlay for 6 metrics |
| F1 Score vs Cost | Scatter plot with regression line |
| Sensitivity vs Specificity Trade-off | Two-panel scatter: per project and aggregated |
| Model Ranking Heatmap | Summary ranking across all metrics |

---

## Advanced: CLI Usage

For users who prefer the command line or need to run the report pipeline on a remote server.

### Generate report from local folders

Place all input files in `input/` then run:

```powershell
# Windows
$env:PYTHONIOENCODING="utf-8"
.\.venv\Scripts\python.exe report\001_report.py

# Custom input/output directories
.\.venv\Scripts\python.exe report\001_report.py --input_dir /path/to/input --output_dir /path/to/output
```

### Generate charts separately

```powershell
# Auto-detect latest data_grafics_*.xlsx in output/
.\.venv\Scripts\python.exe report\graphic.py

# Specify XLSX + custom output directory
.\.venv\Scripts\python.exe report\graphic.py output\data_grafics_20260308.xlsx -o output\my_charts
```

Charts are saved as PNG at 300 DPI in `output/figures_custom/`.

### R charts (optional)

```powershell
Rscript report\graphic.R
```

Requires: `readxl`, `ggplot2`, `dplyr`, `tidyr`, `scales`, `ggrepel`, `fmsb`.

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `CONCURRENT_WORKERS` | `20` | Starting concurrent workers |
| `CONCURRENT_MAX` | `40` | Maximum concurrent workers |
| `CONCURRENT_MIN` | `2` | Minimum concurrent workers |
| `RECORD_MAX_RETRIES` | `3` | Per-record retry attempts |
| `AIUP_AFTER` | `5` | Successes before adding +1 worker |
| `OPENAI_MAX_RETRIES` | `5` | Max retries per API call |
| `OPENAI_BASE_BACKOFF` | `1.0` | Base exponential backoff (seconds) |

---

## Kappa Interpretation (Landis & Koch, 1977)

| Kappa | Agreement |
|-------|-----------|
| < 0 | Poor |
| 0.00–0.20 | Slight |
| 0.21–0.40 | Fair |
| 0.41–0.60 | Moderate |
| 0.61–0.80 | Substantial |
| 0.81–1.00 | Almost Perfect |

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `python` is not recognized | Python not on PATH — reinstall Python and check "Add Python to PATH" on the first installer screen |
| `.venv\Scripts\pip.exe` not found | Run `python -m venv .venv` from the project root |
| `UnicodeEncodeError` on Windows (CLI) | Run `$env:PYTHONIOENCODING="utf-8"` in the same terminal before the report command |
| `No module named 'docx'` | Run `.\.venv\Scripts\pip.exe install python-docx` |
| `No module named 'matplotlib'` | Run `.\.venv\Scripts\pip.exe install matplotlib` |
| `No module named 'fastapi'` | Run `.\.venv\Scripts\pip.exe install -r requirements.txt` |
| `Could not import module "backend"` | Run uvicorn from the **project root** (`AI-tools/`), not from inside `scripts/` |
| Report finds no files | Check that all files follow the exact naming conventions above (case-sensitive project names) |
| Metadata mismatch warnings | `project`, `code`, and `model` in `metadata.xlsx` must match AI result filenames exactly |
| XLSX library not loaded | Access the app via http://localhost:8000, not by opening `index.html` directly |
| Charts not generated | Ensure `matplotlib` and `numpy` are installed (`pip install -r requirements.txt`) |
| Title matching fails | Use the same source TIAB spreadsheet for both human review and AI screening to guarantee identical titles |

---

## License
