import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { LegacyWorkspaceFrame } from '@/legacy/LegacyWorkspaceFrame';

export interface IAewttrWorkspaceAppProps {
  readonly context: WebPartContext;
  readonly description: string;
  readonly isDarkTheme: boolean;
}

export function AewttrWorkspaceApp(): React.ReactElement {
  return (
    <div style={{ margin: 0, padding: 0 }}>
      <LegacyWorkspaceFrame />
    </div>
  );
}
