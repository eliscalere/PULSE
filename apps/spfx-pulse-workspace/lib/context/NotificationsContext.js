import * as React from 'react';
export var NotificationsContext = React.createContext(undefined);
export function useNotifications() {
    var context = React.useContext(NotificationsContext);
    if (!context) {
        throw new Error('NotificationsContext is not available.');
    }
    return context;
}
//# sourceMappingURL=NotificationsContext.js.map