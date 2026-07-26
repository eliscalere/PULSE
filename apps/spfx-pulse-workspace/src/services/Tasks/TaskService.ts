import { SPFI } from '@pnp/sp';

import { LIST_KEYS } from '@/constants/sharePointResources';
import { ITaskService } from '@/interfaces/IAppServices';
import { TaskItem } from '@/models/TaskItem';
import { SharePointListService } from '@/services/SharePoint/SharePointListService';
import { AewttrWorkspaceConfig } from '@/config/AewttrWorkspaceConfig';

export class TaskService implements ITaskService {
  private readonly lists: SharePointListService;

  public constructor(sp: SPFI) {
    this.lists = new SharePointListService(sp);
  }

  public async getTasksByProject(projectId: number): Promise<readonly TaskItem[]> {
    const items = await this.lists.getItems<Record<string, unknown>>(AewttrWorkspaceConfig.lists[LIST_KEYS.tasks], [
      'Id',
      'Title',
      'RelatedProjectId',
      'TaskStatus',
      'PercentComplete',
      'DueDate'
    ]);

    return items
      .filter((item) => Number(item.RelatedProjectId) === projectId)
      .map((item) => ({
        id: Number(item.Id),
        title: String(item.Title),
        projectId: Number(item.RelatedProjectId),
        status: (item.TaskStatus as TaskItem['status']) ?? 'To Do',
        percentComplete: Number(item.PercentComplete ?? 0),
        priority: 'Medium',
        dueDate: item.DueDate ? String(item.DueDate) : undefined
      }));
  }
}
