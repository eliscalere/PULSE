import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

import { APP_NAME } from '@/constants/appConstants';
import { PageHeader } from '@/components/common/PageHeader';
import { AppNavigation } from '@/components/navigation/AppNavigation';
import { ToastRegion } from '@/components/common/ToastRegion';
import { AppRoute } from '@/types/AppRoute';

interface AppShellProps {
  readonly activeRoute: AppRoute;
  readonly onNavigate: (route: AppRoute) => void;
  readonly children: React.ReactNode;
}

export function AppShell(props: AppShellProps): React.ReactElement {
  return (
    <FluentProvider theme={webLightTheme}>
      <div>
        <PageHeader title={APP_NAME} subtitle="Operational SharePoint workspace." />
        <AppNavigation activeRoute={props.activeRoute} onNavigate={props.onNavigate} />
        <ToastRegion />
        <section>{props.children}</section>
      </div>
    </FluentProvider>
  );
}
