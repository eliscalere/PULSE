import { SPFI } from '@pnp/sp';

import { LIST_KEYS } from '@/constants/sharePointResources';
import { IMeetingService } from '@/interfaces/IAppServices';
import { Meeting } from '@/models/Meeting';
import { SharePointListService } from '@/services/SharePoint/SharePointListService';
import { AewttrWorkspaceConfig } from '@/config/AewttrWorkspaceConfig';

export class MeetingService implements IMeetingService {
  private readonly lists: SharePointListService;

  public constructor(sp: SPFI) {
    this.lists = new SharePointListService(sp);
  }

  public async getRecentMeetings(): Promise<readonly Meeting[]> {
    const items = await this.lists.getItems<Record<string, unknown>>(AewttrWorkspaceConfig.lists[LIST_KEYS.meetings], [
      'Id',
      'Title',
      'MeetingDate',
      'Summary'
    ]);

    return items.map((item) => ({
      id: Number(item.Id),
      title: String(item.Title),
      meetingDate: String(item.MeetingDate ?? ''),
      summary: item.Summary ? String(item.Summary) : undefined
    }));
  }
}
