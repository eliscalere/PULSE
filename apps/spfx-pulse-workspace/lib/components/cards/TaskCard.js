import * as React from 'react';
import { Card, Body1, Subtitle2 } from '@fluentui/react-components';
import { toDisplayDate } from '@/utils/dateUtils';
export function TaskCard(props) {
    return (React.createElement(Card, null,
        React.createElement(Subtitle2, null, props.task.title),
        React.createElement(Body1, null, props.task.status),
        React.createElement(Body1, null, toDisplayDate(props.task.dueDate))));
}
//# sourceMappingURL=TaskCard.js.map