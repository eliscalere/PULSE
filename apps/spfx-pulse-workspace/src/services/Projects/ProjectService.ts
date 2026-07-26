import { SPFI } from '@pnp/sp';

import { LIST_KEYS } from '@/constants/sharePointResources';
import { IProjectService } from '@/interfaces/IAppServices';
import { Project } from '@/models/Project';
import { SharePointListService } from '@/services/SharePoint/SharePointListService';
import { AewttrWorkspaceConfig } from '@/config/AewttrWorkspaceConfig';

export class ProjectService implements IProjectService {
  private readonly lists: SharePointListService;

  public constructor(sp: SPFI) {
    this.lists = new SharePointListService(sp);
  }

  public async getProjects(): Promise<readonly Project[]> {
    const items = await this.lists.getItems<Record<string, unknown>>(AewttrWorkspaceConfig.lists[LIST_KEYS.projects], [
      'Id',
      'Title',
      'ProjectCode',
      'ProjectHealth',
      'ProjectStatus',
      'TargetDate',
      'TeamName'
    ]);

    return items.map((item) => ({
      id: Number(item.Id),
      title: String(item.Title),
      code: String(item.ProjectCode ?? ''),
      description: '',
      health: (item.ProjectHealth as Project['health']) ?? 'Blue',
      status: (item.ProjectStatus as Project['status']) ?? 'Not Started',
      priority: 'Medium',
      targetDate: item.TargetDate ? String(item.TargetDate) : undefined,
      team: item.TeamName ? String(item.TeamName) : undefined
    }));
  }

  public async getProjectById(id: number): Promise<Project | undefined> {
    const projects = await this.getProjects();
    return projects.find((project) => project.id === id);
  }
}
