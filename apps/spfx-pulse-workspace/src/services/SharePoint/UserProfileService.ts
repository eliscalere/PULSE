import { SPFI } from '@pnp/sp';

import { User } from '@/models/User';

export class UserProfileService {
  public constructor(private readonly sp: SPFI) {}

  public async getCurrentUser(): Promise<User> {
    const user = await this.sp.web.currentUser.select('Id', 'Title', 'Email')();
    const groups = await this.sp.web.currentUser.groups.select('Title')();
    const groupTitles = groups.map((group) => group.Title);

    return {
      id: user.Id,
      title: user.Title,
      email: user.Email,
      role: groupTitles[0] ?? 'Member',
      isSiteOwner: groupTitles.some((title) => title.includes('Owners')),
      isSiteMember: groupTitles.some((title) => title.includes('Members')),
      isAdministrator: groupTitles.some((title) => title.includes('Owners') || title.includes('Administrators'))
    };
  }
}
