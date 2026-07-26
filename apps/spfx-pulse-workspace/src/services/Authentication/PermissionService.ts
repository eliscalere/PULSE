import { SPFI } from '@pnp/sp';

import { IPermissionService } from '@/interfaces/IAppServices';
import { User } from '@/models/User';
import { UserProfileService } from '@/services/SharePoint/UserProfileService';

export class PermissionService implements IPermissionService {
  private readonly profiles: UserProfileService;

  public constructor(sp: SPFI) {
    this.profiles = new UserProfileService(sp);
  }

  public async getCurrentUser(): Promise<User> {
    return this.profiles.getCurrentUser();
  }

  public async isAdministrator(): Promise<boolean> {
    const user = await this.getCurrentUser();
    return user.isAdministrator;
  }
}
