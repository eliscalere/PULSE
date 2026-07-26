import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { APP_NAME } from '@/constants/appConstants';
import { PageHeader } from '@/components/common/PageHeader';
import { AppNavigation } from '@/components/navigation/AppNavigation';
import { ToastRegion } from '@/components/common/ToastRegion';
export function AppShell(props) {
    return (React.createElement(FluentProvider, { theme: webLightTheme },
        React.createElement("div", null,
            React.createElement(PageHeader, { title: APP_NAME, subtitle: "Operational SharePoint workspace." }),
            React.createElement(AppNavigation, { activeRoute: props.activeRoute, onNavigate: props.onNavigate }),
            React.createElement(ToastRegion, null),
            React.createElement("section", null, props.children))));
}
//# sourceMappingURL=AppShell.js.map