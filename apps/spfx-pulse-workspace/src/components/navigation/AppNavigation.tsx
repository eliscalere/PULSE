import * as React from 'react';
import { Button } from '@fluentui/react-components';

import { NAVIGATION_ITEMS } from '@/constants/navigation';
import { AppRoute } from '@/types/AppRoute';

interface AppNavigationProps {
  readonly activeRoute: AppRoute;
  readonly onNavigate: (route: AppRoute) => void;
}

export function AppNavigation(props: AppNavigationProps): React.ReactElement {
  return (
    <nav>
      {NAVIGATION_ITEMS.map((item) => (
        <Button
          key={item.key}
          appearance={item.key === props.activeRoute ? 'primary' : 'secondary'}
          onClick={() => props.onNavigate(item.key)}
        >
          {item.label}
        </Button>
      ))}
    </nav>
  );
}
