# SRID COM - Project Steering

## 1) Purpose
This document is the single source of truth for the SRID COM project status:
- functional goals
- technical decisions
- what is done
- what remains
- deployment and validation rules

## 2) Product Goals
- Provide a clear dashboard for business and management decisions.
- Keep SRID and SRID Genetics visibility explicit.
- Ensure stable deployment on cPanel under subpaths.
- Keep data quality high (bank names normalization, reliable filters).
- Keep UX simple, predictable, and visually coherent.

## 3) Current Technical Stack
- Backend: Flask + SQLAlchemy
- Database: SQLite
- Frontend: Jinja templates + HTMX + Chart.js + DaisyUI/Tailwind
- PWA: manifest + service worker
- Hosting: cPanel/Passenger

## 4) Decisions (Validated)

### 4.1 Dashboard Filters
- Year filters use custom opaque dropdown style for consistency and compactness.
- Month filter in Top 5 clients also uses the same custom dropdown style.
- Opaque dropdown style tokens currently used:
  - background: #0f172a
  - text: #f8fafc
  - border: #334155
  - hover/selected: #1e293b
  - gray scrollbar

### 4.2 Status Monthly Chart
- "Montants par statut - par mois" now supports 3 views:
  - total
  - srid
  - genetics
- Year filter remains active and updates selected view data.

### 4.3 Data Quality (Bank Names)
- Bank values are normalized and cleaned.
- Create/edit/import paths normalize bank names before persistence.
- Strict suggestion behavior preferred over fuzzy browser datalist behavior.

### 4.4 Auth and Case Handling
- Username/login handling normalized to lowercase.

### 4.5 PWA and Subpath Deployment
- Service worker route aligned for subpath deployment.
- Static asset path handling corrected to avoid root path breakage.

## 5) Implemented Scope (Done)
- Dashboard with KPI blocks and multiple charts.
- Year filter restoration to original behavior where requested.
- Top 5 clients year/month filtering.
- Monthly status chart with view selector (total/srid/genetics).
- Bank normalization helpers and cleaned database values.
- PWA fixes (scope/path consistency).
- cPanel/Passenger deployment fixes and rewrite handling guidance.

## 6) Open Items / Backlog
- Label consistency task (pending): rename visible "Genetics" label to "SRID Genetics" where required.
- Optional UI pass: final polish of dropdown/button alignment for all cards.
- Optional analytics extensions:
  - export filtered chart data
  - additional management KPIs

## 7) Operational Rules
- Do not remove existing filter behavior unless explicitly requested.
- Keep data filters deterministic and user-predictable.
- Prefer minimal, targeted UI changes to avoid regressions.
- Validate template compilation after dashboard edits.
- Validate Python syntax after backend/API edits.

## 8) API Notes (Dashboard)
- /api/dashboard/monthly
- /api/dashboard/societes
- /api/dashboard/statuts-monthly
  - now returns total and view-based payloads (total/srid/genetics)
- /api/dashboard/top-clients

## 9) Validation Checklist (Before Release)
- Dashboard page renders without template errors.
- Year/month filters update data correctly.
- Status chart switches correctly between total/srid/genetics.
- Top 5 updates correctly on year/month change.
- Dropdown custom style remains opaque and readable.
- PWA manifest and service worker behavior verified on target domain/subpath.

### 10) Change Log (High Level)
- Deployment hardening for cPanel subpath.
- Dashboard filter behavior restoration and refinement.
- Opaque custom dropdown style introduction.
- Month dropdown aligned to same custom style.
- Status chart upgraded to 3 data views.
- Data normalization for banks and stricter suggestion UX.

---
Owner: SRID project team
Doc type: Steering / project reference
Last update: 2026-06-21
