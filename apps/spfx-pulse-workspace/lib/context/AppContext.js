import * as React from 'react';
export var AppContext = React.createContext(undefined);
export function useAppContext() {
    var context = React.useContext(AppContext);
    if (!context) {
        throw new Error('AppContext is not available.');
    }
    return context;
}
//# sourceMappingURL=AppContext.js.map