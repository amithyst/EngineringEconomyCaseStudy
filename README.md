# Engineering Economy Case Study

This repository is the team workspace for the Engineering Economy case study assignment.

## Goal

Build a defendable investment recommendation for the Ningbo Yifenghui shopping mall acquisition case by:

- constructing the annual cash flow model for 2026-2029
- estimating the reasonable acquisition price under `MARR = 12%`
- performing sensitivity analysis on leasing assumptions
- writing the non-financial risk and operating strategy discussion

## Repository Layout

- `CaseStudy作业要求/`: original assignment folder kept at repository root for direct access
- `docs/assignment/`: original assignment prompt and reference materials
- `docs/project/`: project management, branch rules, and task breakdown
- `docs/analysis/`: working notes, model logic, and writeup drafts
- `data/raw/`: untouched original source files
- `data/processed/`: cleaned and derived data
- `src/`: scripts used to extract data or generate analysis artifacts
- `outputs/`: generated tables, charts, and deliverables

## Branch Strategy

- `main`: integration branch for the final submission version
- `codex`: Codex working branch
- `cc`: CC working branch
- `libowen-integration`: optional branch for Li Bowen's agent to merge and reconcile content before `main`

## Working Rules

1. Keep raw files unchanged in `data/raw/`.
2. Put all reusable logic in `src/`, not inside ad hoc spreadsheets only.
3. Record major assumptions and decisions in `docs/analysis/`.
4. Merge into `main` only after the branch output is reviewable.

## Current Inputs

- Root assignment folder: `CaseStudy作业要求/`
- Assignment brief: `docs/assignment/EE2025 Case Study.docx`
- Source workbook: `data/raw/EE2025 Case Study Sheet.xlsm`
- Original archive: `data/raw/CaseStudy.zip`
