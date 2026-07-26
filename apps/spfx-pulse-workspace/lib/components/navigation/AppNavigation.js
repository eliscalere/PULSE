import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { NAVIGATION_ITEMS } from '@/constants/navigation';
export function AppNavigation(props) {
    return (React.createElement("nav", null, NAVIGATION_ITEMS.map(function (item) { return (React.createElement(Button, { key: item.key, appearance: item.key === props.activeRoute ? 'primary' : 'secondary', onClick: function () { return props.onNavigate(item.key); } }, item.label)); })));
}
//# sourceMappingURL=AppNavigation.js.map