# Development Guide

## How to add a new page

1. Create a folder under `src/pages/<PageName>`.
2. Add a functional component.
3. Register the route in `src/types/AppRoute.ts`.
4. Add a navigation item in `src/constants/navigation.ts`.
5. Extend the page switch in `src/webparts/aewttrWorkspace/components/AewttrWorkspaceApp.tsx`.

## How to add a new SharePoint list

1. Add the list definition to `src/schema/lists.json`.
2. Add any seed records to `src/schema/seedData.json`.
3. Add the list key to `src/constants/sharePointResources.ts`.
4. Add its configured title to `src/config/AewttrWorkspaceConfig.ts`.
5. Create or extend a service to read and write that list.

## How to add a new service

1. Create the service under the appropriate `src/services/*` folder.
2. Add its interface to `src/interfaces/IAppServices.ts` if it should be app-wide.
3. Register it in `useAppInitialization`.
4. Consume it through hooks or page-level composition, never directly from deeply nested components.

## Component hierarchy guidance

- Pages compose sections, cards, tables, and dialogs.
- Reusable controls stay under `src/components`.
- Business logic should remain in hooks or services.
- SharePoint calls should remain in services only.
