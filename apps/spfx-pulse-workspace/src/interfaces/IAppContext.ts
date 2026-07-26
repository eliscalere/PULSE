import { WebPartContext } from '@microsoft/sp-webpart-base';

import { AppRoute } from '@/types/AppRoute';
import { IAppServices } from './IAppServices';
import { User } from '@/models/User';
import { Notification } from '@/models/Notification';

export interface IAppContext {
  readonly context: WebPartContext;
  readonly services: IAppServices;
  readonly currentUser?: User;
  readonly activeRoute: AppRoute;
  readonly notifications: readonly Notification[];
  readonly isInitializing: boolean;
  readonly error?: Error;
}
