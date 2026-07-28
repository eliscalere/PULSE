# AEWTTR Workspace SPFx Solution

This project re-architects the legacy `PULSE` prototype into a production-oriented SharePoint Framework solution using React, TypeScript, Fluent UI, and PnPjs.

## What is included

- SPFx `1.23.0` solution scaffold
- React `17.0.1` and TypeScript enterprise application shell
- Schema-driven SharePoint provisioning
- Typed service layer for SharePoint lists and libraries
- Context and hook-based UI composition
- Initial page structure for dashboard, projects, meetings, tasks, documents, reports, settings, and admin
- Centralized error handling and notification infrastructure

## Solution structure

The project is organized around long-term maintainability:

- `src/components`: reusable UI elements, layout, dialogs, tables, cards, and charts
- `src/pages`: route-level page composition
- `src/services`: business logic and SharePoint data access
- `src/hooks`: reusable async and initialization hooks
- `src/context`: global application state providers
- `src/models`, `src/interfaces`, `src/types`: strong typing and contracts
- `src/schema`: list, library, view, and seed-data definitions
- `src/provisioning`: validation, creation, seeding, and upgrade plumbing
- `src/config`, `src/permissions`, `src/telemetry`: cross-cutting concerns

## Local setup

1. Use the bundled local Node `22.x` runtime in `.local/node22`.
2. Run `./use-node22.sh npm install`.
3. Run `./use-node22.sh gulp serve`.
4. Use the configured SharePoint workbench page or update `config/serve.json` with your site URL.

## Important note

This workspace includes a local Node `22.17.0` runtime under `.local/node22` so you can test without changing your system Node installation.

## Migration guidance

The legacy prototype remains in `../PULSE` as a feature and UX reference. The new codebase is intentionally not a wrapper around that prototype. Instead, it establishes:

- a typed domain model
- a service boundary around SharePoint
- a provisioning pipeline that can create missing resources
- a composable page architecture for phased feature implementation
