# Architecture

## Overview

SharePoint Online is treated as the system of record. The SPFx web part is the presentation host, while PnPjs-backed services encapsulate all list, library, and user-profile operations.

## Layers

1. `webparts/aewttrWorkspace`
   Hosts the SPFx entry point and injects `WebPartContext`.
2. `components` and `pages`
   Render Fluent UI experiences with functional components only.
3. `hooks` and `context`
   Manage initialization, route state, notifications, and user/session state.
4. `services`
   Own all SharePoint communication and business operations.
5. `schema` and `provisioning`
   Define and validate the backend footprint without hardcoded list IDs.
6. `telemetry` and `permissions`
   Centralize cross-cutting behavior.

## Data flow

1. SPFx loads the web part.
2. `useAppInitialization` creates the service registry.
3. `PermissionService` resolves the current user.
4. `ProvisioningService` validates and optionally creates required SharePoint resources.
5. Pages call typed services through hooks and context.

## Provisioning approach

- Lists and libraries are defined in JSON schema.
- Provisioners create missing resources only.
- Seed data is inserted idempotently.
- Existing data is never deleted.
- Upgrade hooks are reserved for future schema evolution.

## Extension path

To add a new functional area:

1. Add models and interfaces.
2. Add or extend schema definitions.
3. Create a service under `src/services`.
4. Add a page under `src/pages`.
5. Compose reusable pieces from `src/components`.
