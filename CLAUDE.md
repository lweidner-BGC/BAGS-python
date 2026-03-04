# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**BAGS (Bedload Analysis and Gravel Sediment Transport)** — a VBA-based Excel/LibreOffice Calc application for calculating bedload transport in gravel-bed streams. It implements multiple sediment transport equations from academic literature (Parker 1990, Parker & Klingeman 1982, Wilcock & Crowe 2003, etc.).

- Language: VBA (Visual Basic for Applications)
- No build system, package manager, or test framework — it's a standalone spreadsheet macro application
- Source code lives in `src_VBA/` as `.bas` files (VBA modules) and `.xdl` files (LibreOffice dialog XML)

## Development

There are no CLI build or test commands. Development requires:
1. Opening the Excel/Calc workbook file
2. Accessing the VBA IDE (Alt+F11 in Excel; Tools > Macros > Edit Macros in LibreOffice)
3. Editing `.bas` modules directly or importing them into the IDE

The `.bas` files in `src_VBA/` are standalone VBA module exports that can be imported into the workbook.

## Architecture

### Application Flow
```
Workbook_Open() → Auto_Open() [BAGSModule1.bas]
  → Welcome Sheet display
  → Agreement acceptance
  → ufEquations dialog (equation selection)
  → Input form dialogs (discharge, slope, grain size, etc.)
  → Equation module calculations
  → Results + BAGSgRAPH charting
```

### Module Responsibilities

| File | Role |
|------|------|
| `BAGSModule1.bas` | Main controller: `Auto_Open()`, `RunSoftware()`, equation dispatch |
| `BAGSModule2.bas` | Utilities: menu management, sheet protection, messaging |
| `BAGSModule3.bas` | Numerical processing: calls into equation-specific modules |
| `Parker90Module.bas` | Parker (1990) surface-based equation |
| `PK82Bakke99Module.bas` | Parker-Klingeman (1982) and Bakke et al. (1999) equations |
| `PKM82D50Module.bas` | Parker-Klingeman-McLean D50 variant |
| `Wilcock01Module.bas` | Wilcock (2001) two-fraction model |
| `Wilcock03Module.bas` | Wilcock & Crowe (2003) mixed-size model |
| `BAGSgRAPH.bas` | Chart/graph generation |
| `ufEquations.bas` | UI: equation selection dialog |
| `uf*.bas` | UI: input dialogs (project, slope, discharge, Manning's n, grain fractions) |

### Data Layer
Calculations use worksheet cells for intermediate storage:
- **Input** sheet: user-entered hydraulic and sediment data
- **Storage** sheet: temporary calculation workspace
- **cp** sheet: lookup tables for omega/sigma parameters

### Key Global State (BAGSModule1.bas)
Boolean flags control which equations are active: `Parker90`, `Parker82`, `PK82`, `Wilcock`, `Wilcock03`, `Bakke`. `OnErrorOn` controls global error handling behavior.

### Key Constants
- Default water density: 1000 kg/m³
- Relative sediment density: 1.65
- Version: 2008.11
