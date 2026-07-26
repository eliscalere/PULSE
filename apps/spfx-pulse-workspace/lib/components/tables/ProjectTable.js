import * as React from 'react';
import { Table, TableBody, TableCell, TableHeader, TableHeaderCell, TableRow } from '@fluentui/react-components';
export function ProjectTable(props) {
    return (React.createElement(Table, null,
        React.createElement(TableHeader, null,
            React.createElement(TableRow, null,
                React.createElement(TableHeaderCell, null, "Project"),
                React.createElement(TableHeaderCell, null, "Code"),
                React.createElement(TableHeaderCell, null, "Status"),
                React.createElement(TableHeaderCell, null, "Health"))),
        React.createElement(TableBody, null, props.projects.map(function (project) { return (React.createElement(TableRow, { key: project.id },
            React.createElement(TableCell, null, project.title),
            React.createElement(TableCell, null, project.code),
            React.createElement(TableCell, null, project.status),
            React.createElement(TableCell, null, project.health))); }))));
}
//# sourceMappingURL=ProjectTable.js.map