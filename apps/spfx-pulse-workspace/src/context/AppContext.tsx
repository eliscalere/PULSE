import * as React from 'react';

import { IAppContext } from '@/interfaces/IAppContext';

export const AppContext = React.createContext<IAppContext | undefined>(undefined);

export function useAppContext(): IAppContext {
  const context = React.useContext(AppContext);

  if (!context) {
    throw new Error('AppContext is not available.');
  }

  return context;
}
