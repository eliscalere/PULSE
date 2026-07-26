import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';
export function ProjectFilterPanel(props) {
    return (React.createElement(Dropdown, { placeholder: "Filter by team", onOptionSelect: function (_, data) { var _a; return props.onTeamChanged(String((_a = data.optionValue) !== null && _a !== void 0 ? _a : '')); } },
        React.createElement(Option, { value: "" }, "All Teams"),
        React.createElement(Option, { value: "PM" }, "PM"),
        React.createElement(Option, { value: "Engineering" }, "Engineering"),
        React.createElement(Option, { value: "Cyber" }, "Cyber"),
        React.createElement(Option, { value: "Support" }, "Support")));
}
//# sourceMappingURL=ProjectFilterPanel.js.map