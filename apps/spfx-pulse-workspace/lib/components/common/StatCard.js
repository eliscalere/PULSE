import * as React from 'react';
import { Card, CardHeader, Body1, Subtitle2 } from '@fluentui/react-components';
export function StatCard(props) {
    return (React.createElement(Card, null,
        React.createElement(CardHeader, { header: React.createElement(Subtitle2, null, props.label) }),
        React.createElement(Body1, null, props.value),
        React.createElement(Body1, null, props.description)));
}
//# sourceMappingURL=StatCard.js.map