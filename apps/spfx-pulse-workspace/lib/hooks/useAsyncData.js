import * as React from 'react';
import { toError } from '@/utils/errorUtils';
export function useAsyncData(load, dependencies) {
    var _a = React.useState({ isLoading: true }), state = _a[0], setState = _a[1];
    React.useEffect(function () {
        var isMounted = true;
        setState({ isLoading: true });
        void load()
            .then(function (data) {
            if (isMounted) {
                setState({ data: data, isLoading: false });
            }
        })
            .catch(function (error) {
            if (isMounted) {
                setState({ error: toError(error), isLoading: false });
            }
        });
        return function () {
            isMounted = false;
        };
    }, dependencies);
    return state;
}
//# sourceMappingURL=useAsyncData.js.map