<p align="center">
  <img src="images/header.svg" alt="DSM Report Comparator overview" width="100%">
</p>

<p align="center">
  <img alt="version" src="https://img.shields.io/badge/version-2.0-0b5cad">
  <img alt="python" src="https://img.shields.io/badge/python-3.10%2B-3776ab">
  <img alt="input" src="https://img.shields.io/badge/input-read--only-0f8a5f">
  <img alt="output" src="https://img.shields.io/badge/output-HTML%20%2B%20Excel-b86e00">
</p>

> Automated DSM report comparison tool for identifying diagnostic changes between two PVER releases.

DSM Report Comparator compares DSM reports from two PVER releases and generates a detailed diagnostics delta report. It reduces manual review effort during software releases, regression analysis, and diagnostics impact assessments.

## What it detects

| | Change | Description |
|---|---|---|
| 🟢 | **Added** | Diagnostics present only in the target release |
| 🔴 | **Removed** | Diagnostics present only in the baseline release |
| 🟡 | **Modified** | FID changes, DFC mapping changes, description and configuration updates |

## How it works

```mermaid
flowchart LR
    A["Base report<br/>PVER n"] --> C{{"Comparator<br/>DFC · FID · DSQ · DFQ"}}
    B["Target report<br/>PVER n+1"] --> C
    C --> D["🟢 Added"]
    C --> E["🟡 Modified"]
    C --> F["🔴 Removed"]
    D --> G["HTML / Excel report"]
    E --> G
    F --> G
    classDef src fill:#12304a,stroke:#2f5578,color:#eaf1f8
    classDef core fill:#0b5cad,stroke:#0b5cad,color:#fff
    classDef add fill:#0f3a2c,stroke:#3fd19a,color:#a7e8cd
    classDef mod fill:#41300e,stroke:#f0b04a,color:#f5d596
    classDef rem fill:#431b1b,stroke:#ff7b76,color:#ffc2bf
    class A,B src
    class C core
    class D add
    class E mod
    class F rem
```

## Prerequisites

> [!NOTE]
> Python 3.10 or later and your DSM report files (`.txt`) are required.

| Required | Recommended |
|---|---|
| Python 3.10+ | Microsoft Excel (for `.xlsx` reports) |
| DSM report files | Chrome, Edge, or Firefox (for HTML reports) |

Verify your Python installation:

```bash
python --version
# Python 3.x.x
```

## Quick start

```mermaid
flowchart LR
    S1["1. Select<br/>baseline"] --> S2["2. Select<br/>target"] --> S3["3. Choose<br/>format"] --> S4["4. Run<br/>comparison"] --> S5["5. Review<br/>report"]
```

1. Select the baseline DSM report.
2. Select the target DSM report.
3. Choose the output format: **HTML**, **Excel**, or **HTML + Excel**.
4. Run the comparison.
5. Review the generated diagnostics comparison report.

## Features

### Case-insensitive prefix support

Prefixes are matched in any letter case. Original entity names are preserved, so comparison accuracy is maintained across DSM generator variants.

| Entity | Example matches |
|---|---|
| DFC | `DFC_` `dfc_` `DfC_` |
| FID | `FID_` `FId_` `fid_` |
| DSQ | `DSQ_` `dsq_` |
| DFQ | `DFQ_` `dfq_` |

Example: `FId_FuHiPGPCtrlDeCatHeatgMin`

### Intelligent FID change detection

A FID is reported as **MODIFIED** when its DFC association changes, even if the description, scheduling parameters, and IUMPR fields stay the same.

```mermaid
flowchart LR
    F["FId_FuHiPGP..."] -- Base --> D1["DFC_A"]
    F -. Target .-> D2["DFC_B"]
    D2 --> M["🟡 FID reported as MODIFIED"]
```

> [!TIP]
> FID-to-DFC mapping changes are also reported separately in the mapping analysis section for full traceability.

### Description handling

- Generated metadata such as `LongName_EN` is removed from descriptions.
- Empty or missing descriptions are shown as `Not Available`.

## Input and output

> [!IMPORTANT]
> Source DSM reports are never modified. All analysis runs in read-only mode and output files are written separately.

Input files use the `.txt` format, for example `Base_dsm_report.txt` and `Target_dsm_report.txt`.

Reports are saved in an automatically created `Output` subfolder inside the folder selected in the **Output** field. If the field is empty, the target report's folder is used as the parent. The **Open Output Folder** button opens this subfolder.

| Format | File |
|---|---|
| HTML | `Comparison_Report.html` |
| Excel | `Comparison_Report.xlsx` |
| HTML + Excel | Both files |

<details>
<summary><b>Project structure</b></summary>

```text
DSM-Report-Comparator/
├── DSM_Comparator.py
├── DSM_Comparator.pyw   (to be executed with python.exe)
├── README.md
├── docs/
│   └── hero.svg
├── Input/
│   ├── dsm_report_1.txt
│   └── dsm_report_2.txt
├── Output/
│   ├── Comparison_Report.html
│   └── Comparison_Report.xlsx
└── logs/
```

</details>

<details>
<summary><b>What's new in 2.0</b></summary>

| Type | Change |
|---|---|
| 🟢 Added | Case-insensitive handling of DFC, FID, DSQ, and DFQ prefixes |
| 🟢 Added | Support for mixed-case FID identifiers (`FID_`, `FId_`, etc.) |
| 🟢 Added | FID change detection based on DFC mapping modifications |
| 🔵 Improved | Automatic removal of metadata suffixes such as `LongName_EN` |
| 🔵 Improved | Consistent handling of empty or NULL descriptions using `Not Available` |
| 🟠 Fixed | Missing FID detection caused by mixed-case prefixes |
| 🟠 Fixed | FID modifications not reported when only DFC associations changed |

</details>

## Benefits

- ✅ Faster diagnostics review
- ✅ Reduced manual comparison effort
- ✅ Improved release analysis
- ✅ Better traceability of diagnostic changes
- ✅ Consistent reporting across DSM generator versions
- ✅ Read-only processing of source DSM reports

## License

RB Internal Engineering Tool.
