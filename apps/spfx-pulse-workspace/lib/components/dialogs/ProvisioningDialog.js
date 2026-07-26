import * as React from 'react';
import { Dialog, DialogBody, DialogContent, DialogSurface, DialogTitle, MessageBar, MessageBarBody } from '@fluentui/react-components';
export function ProvisioningDialog(props) {
    if (!props.status || props.status.isReady) {
        return React.createElement(React.Fragment, null);
    }
    return (React.createElement(Dialog, { open: true },
        React.createElement(DialogSurface, null,
            React.createElement(DialogBody, null,
                React.createElement(DialogTitle, null, "SharePoint setup required"),
                React.createElement(DialogContent, null, props.status.messages.map(function (message, index) { return (React.createElement(MessageBar, { key: index, intent: message.level === 'error' ? 'error' : 'warning' },
                    React.createElement(MessageBarBody, null, message.text))); }))))));
}
//# sourceMappingURL=ProvisioningDialog.js.map