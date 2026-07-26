import * as React from 'react';
import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';
import { useNotifications } from '@/context/NotificationsContext';
export function ToastRegion() {
    var _a = useNotifications(), notifications = _a.notifications, dismiss = _a.dismiss;
    return (React.createElement(React.Fragment, null, notifications.map(function (notification) { return (React.createElement(MessageBar, { key: notification.id, intent: notification.intent },
        React.createElement(MessageBarBody, null,
            React.createElement(MessageBarTitle, null, notification.title),
            notification.message),
        React.createElement("button", { onClick: function () { return dismiss(notification.id); } }, "Dismiss"))); })));
}
//# sourceMappingURL=ToastRegion.js.map