import * as React from 'react';
import { Card, Body1, Subtitle2 } from '@fluentui/react-components';
import { toDisplayDate } from '@/utils/dateUtils';
export function ProjectCard(props) {
    var project = props.project;
    return (React.createElement(Card, null,
        React.createElement(Subtitle2, null, project.title),
        React.createElement(Body1, null, project.code),
        React.createElement(Body1, null, project.health),
        React.createElement(Body1, null, toDisplayDate(project.targetDate))));
}
//# sourceMappingURL=ProjectCard.js.map