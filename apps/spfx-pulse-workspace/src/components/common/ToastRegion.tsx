import * as React from 'react';
import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';

import { useNotifications } from '@/context/NotificationsContext';

export function ToastRegion(): React.ReactElement {
  const { notifications, dismiss } = useNotifications();

  return (
    <>
      {notifications.map((notification) => (
        <MessageBar key={notification.id} intent={notification.intent}>
          <MessageBarBody>
            <MessageBarTitle>{notification.title}</MessageBarTitle>
            {notification.message}
          </MessageBarBody>
          <button onClick={() => dismiss(notification.id)}>Dismiss</button>
        </MessageBar>
      ))}
    </>
  );
}
