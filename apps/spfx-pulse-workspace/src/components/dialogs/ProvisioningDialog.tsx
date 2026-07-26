import * as React from 'react';
import { Dialog, DialogBody, DialogContent, DialogSurface, DialogTitle, MessageBar, MessageBarBody } from '@fluentui/react-components';

import { ProvisioningStatus } from '@/provisioning/ProvisioningModels';

interface ProvisioningDialogProps {
  readonly status?: ProvisioningStatus;
}

export function ProvisioningDialog(props: ProvisioningDialogProps): React.ReactElement {
  if (!props.status || props.status.isReady) {
    return <></>;
  }

  return (
    <Dialog open>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>SharePoint setup required</DialogTitle>
          <DialogContent>
            {props.status.messages.map((message, index) => (
              <MessageBar key={index} intent={message.level === 'error' ? 'error' : 'warning'}>
                <MessageBarBody>{message.text}</MessageBarBody>
              </MessageBar>
            ))}
          </DialogContent>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
