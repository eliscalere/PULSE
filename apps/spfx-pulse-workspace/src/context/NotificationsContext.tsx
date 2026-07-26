import * as React from 'react';

import { Notification } from '@/models/Notification';

export interface INotificationsContext {
  readonly notifications: readonly Notification[];
  push(notification: Notification): void;
  dismiss(id: string): void;
}

export const NotificationsContext = React.createContext<INotificationsContext | undefined>(undefined);

export function useNotifications(): INotificationsContext {
  const context = React.useContext(NotificationsContext);

  if (!context) {
    throw new Error('NotificationsContext is not available.');
  }

  return context;
}
