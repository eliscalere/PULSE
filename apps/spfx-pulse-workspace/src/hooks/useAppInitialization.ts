import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';

import { IAppServices } from '@/interfaces/IAppServices';
import { User } from '@/models/User';
import { ProvisioningStatus } from '@/provisioning/ProvisioningModels';
import { ProjectService } from '@/services/Projects/ProjectService';
import { TaskService } from '@/services/Tasks/TaskService';
import { MeetingService } from '@/services/Meetings/MeetingService';
import { DocumentService } from '@/services/Documents/DocumentService';
import { PermissionService } from '@/services/Authentication/PermissionService';
import { ProvisioningService } from '@/provisioning/ProvisioningService';
import { SettingsService } from '@/services/SettingsService';
import { PnPjsProvider } from '@/services/SharePoint/PnPjsProvider';
import { toError } from '@/utils/errorUtils';

export interface AppInitializationState {
  readonly services: IAppServices;
  readonly currentUser?: User;
  readonly provisioning?: ProvisioningStatus;
  readonly isInitializing: boolean;
  readonly error?: Error;
}

function createServices(sp: SPFI): IAppServices {
  return {
    projects: new ProjectService(sp),
    tasks: new TaskService(sp),
    meetings: new MeetingService(sp),
    documents: new DocumentService(sp),
    permissions: new PermissionService(sp),
    provisioning: new ProvisioningService(sp),
    settings: new SettingsService()
  };
}

export function useAppInitialization(context: WebPartContext): AppInitializationState {
  const services = React.useMemo(() => createServices(PnPjsProvider.create(context)), [context]);
  const [state, setState] = React.useState<AppInitializationState>({
    services,
    isInitializing: true
  });

  React.useEffect(() => {
    let isMounted = true;

    void (async () => {
      try {
        const [currentUser, provisioning] = await Promise.all([
          services.permissions.getCurrentUser(),
          services.provisioning.validateAndProvision()
        ]);

        if (isMounted) {
          setState({
            services,
            currentUser,
            provisioning,
            isInitializing: false
          });
        }
      } catch (error: unknown) {
        if (isMounted) {
          setState({
            services,
            error: toError(error),
            isInitializing: false
          });
        }
      }
    })();

    return () => {
      isMounted = false;
    };
  }, [services]);

  return state;
}
