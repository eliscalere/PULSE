import * as React from 'react';

import { AsyncState } from '@/types/AsyncState';
import { toError } from '@/utils/errorUtils';

export function useAsyncData<T>(load: () => Promise<T>, dependencies: readonly unknown[]): AsyncState<T> {
  const [state, setState] = React.useState<AsyncState<T>>({ isLoading: true });

  React.useEffect(() => {
    let isMounted = true;
    setState({ isLoading: true });

    void load()
      .then((data) => {
        if (isMounted) {
          setState({ data, isLoading: false });
        }
      })
      .catch((error: unknown) => {
        if (isMounted) {
          setState({ error: toError(error), isLoading: false });
        }
      });

    return () => {
      isMounted = false;
    };
  }, dependencies);

  return state;
}
