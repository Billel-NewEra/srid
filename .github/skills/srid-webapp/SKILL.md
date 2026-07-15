---
name: srid-webapp
description: "Use when working on the SRID Flask web app: consultation table, dashboard, logistique gestion, bons de commande, frais, referentiels, HTMX partials, SQLAlchemy pagination, DaisyUI/Tailwind UI consistency, or SRID-specific UX rules."
---

# SRID Web App

## Use When

Use this skill when you need to modify or review the SRID web application in this workspace, especially for:
- Flask routes in app.py
- SQLAlchemy models in models.py
- Jinja templates under templates/
- HTMX partial refresh flows
- Dashboard, Consultation, Logistique, Bons, Frais, and Referentiels
- Pagination, sorting, filtering, and responsive table behavior
- DaisyUI/Tailwind consistency in the existing UI

## Project Shape

Main files:
- app.py: Flask routes, business logic, filtering, pagination, HTMX partial endpoints
- models.py: SQLAlchemy models
- templates/: full pages and partials
- templates/partials/: reusable HTMX table/list fragments
- static/: front-end assets

Key UI areas:
- consultation.html + partials/operations_table.html
- dashboard.html + partials/dashboard_kpis.html
- logistique_bons.html + partials/logistique_bons_table.html
- logistique_gestion.html + partials/logistique_gestion_table.html
- partials/logistique_frais_table.html
- referentiels.html + referential partial lists
- logistique_referentiels.html + fournisseurs list partial

## Working Rules

Follow these rules when changing SRID:
- Preserve the current Flask + Jinja + HTMX architecture. Do not introduce a front-end framework.
- Prefer minimal, local changes over broad refactors.
- Keep pagination, filters, and sorting state consistent across HTMX refreshes.
- When adding a new filter or per-page control, ensure the value is propagated in:
  - the Flask route/query builder
  - the partial template
  - the page-level HTMX includes or JavaScript helpers
- Keep desktop and mobile renderings aligned when a table has both forms.
- Do not duplicate business logic between full-page routes and HTMX partial routes if a helper can be reused.
- Respect the existing DaisyUI/Tailwind visual language instead of redesigning components.

## SRID-Specific UX Preferences

These preferences were established during recent work:
- Consultation table should stay compact on desktop without ugly horizontal overflow on common laptop screens.
- Long status labels should not overlap action buttons.
- Avoid unnecessary columns when they do not add value to the user.
- In logistique, Frais is an independent view, not a copy of Gestion.
- Left sidebar navigation can drive the active logistique section.
- Notifications for logistique should appear only in the Gestion section, not in Frais.
- Dashboard KPI filters should be embedded naturally inside the card rather than floating awkwardly above it.

## Common Change Patterns

### Add or update paginated tables

When modifying a paginated table:
1. Update the backend query helper or route to accept page and per-page values.
2. Return total, total_pages, current page, and per-page selection to the template.
3. Preserve search/filter/sort/per-page in pagination links or HTMX calls.
4. Validate empty-state colspan or layout after column changes.

### Add a new HTMX filter

When adding a filter:
1. Add the input/select to the page template.
2. Include it in hx-include attributes.
3. Include it in any JavaScript URLSearchParams builder.
4. Apply it in the Flask query.
5. Preserve it during pagination and sorting.

### Adjust table density

When users report a table looks cramped or ugly:
- First remove low-value columns before shrinking everything.
- Prefer compact typography and better column priority over abbreviating critical business text.
- Use horizontal scrolling only when truly necessary.
- Keep visual behavior aligned with the nearest comparable SRID table.

## Validation Checklist

Before finishing:
- Verify app.py still compiles.
- Check templates for errors.
- Confirm HTMX flows still preserve search, sort, filter, and per-page state.
- For layout changes, inspect both desktop and mobile template branches when both exist.
- If a change touches consultation/logistique tables, verify pagination controls and empty states still make sense.

## Suggested Commands

Useful validation command:
```bash
cd /c/WorkSpace_AWS/SRID && python -m py_compile app.py
```

## Files Worth Checking First

If the task mentions:
- consultation: templates/consultation.html, templates/partials/operations_table.html, app.py
- dashboard: templates/dashboard.html, templates/partials/dashboard_kpis.html, app.py
- bons: templates/logistique_bons.html, templates/partials/logistique_bons_table.html, app.py
- gestion or frais: templates/logistique_gestion.html, templates/partials/logistique_gestion_table.html, templates/partials/logistique_frais_table.html, app.py
- referentiels: templates/referentiels.html, templates/logistique_referentiels.html, related partials, app.py
