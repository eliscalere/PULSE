# PULSE Documentation

## Published documents (canonical)

The 8 PDFs in `published/` are the authoritative, unclassified release. Start here.

| # | File | What it covers |
|---|---|---|
| 01 | `Brand_Identity_Guide.pdf` | Mark use, color, typography, tone |
| 02 | `Product_Brief.pdf` | What PULSE is and the problem it solves |
| 03 | `General_Overview.pdf` | Operating model, roles, SharePoint context |
| 04 | `Internal_User_Guide.pdf` | Day-to-day usage for all PULSE users |
| 05 | `Technical_Reference.pdf` | Architecture, schema, admin, browser package |
| 06 | `Standard_Operating_Procedures.pdf` | Eight SOPs (projects, meetings, travel, tickets…) |
| 07 | `Documentation_Verification_and_Governance.pdf` | Controlled-doc governance and traceability |
| 08 | `Codebase_Setup_and_Firepit_Packaging_Guide.pdf` | Source layout, build, packaging, and release |

## Folder map

| Folder | What it is |
|---|---|
| `published/` | **Canonical PDFs — start here** |
| `architecture/` | Deep technical references: SP list schema, implementation notes |
| `diagrams/` | Architecture and data-relationship diagrams (PNG + SVG) |
| `handoff/` | Forge deployment guide, FS-specific steps |
| `requirements/` | Original requirements and implementation notes |
| `samples/` | Sample JSON for the Overview PPTX export |
| `screenshots/` | Annotated UI screenshots |
| `legacy/` | Superseded drafts — kept for audit trail only |
| `current/` | Pre-publication Word/MD drafts — superseded by `published/` |

## Developer entry points

- **App code** → `../assets/js/` and `../assets/css/`
- **Build + package** → `../scripts/build-sharepoint-package.js`
- **Schema reference** → `OVERVIEW-PPTX-DATA-SCHEMA.md` and `samples/`
- **Technical deep-dive** → `published/05_PULSE_Technical_Reference.pdf`
- **Deployment** → `handoff/FS-FORGE-STEPS.md` then `published/08_…_Packaging_Guide.pdf`

## Version

Current release: **2026.07.22**
