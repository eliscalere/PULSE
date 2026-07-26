import { Project } from '@/models/Project';
import { TaskItem } from '@/models/TaskItem';
import { Meeting } from '@/models/Meeting';
import { DocumentRecord } from '@/models/DocumentRecord';
import { User } from '@/models/User';
import { ProvisioningStatus } from '@/provisioning/ProvisioningModels';
import { Configuration } from '@/models/Configuration';

export interface IProjectService {
  getProjects(): Promise<readonly Project[]>;
  getProjectById(id: number): Promise<Project | undefined>;
}

export interface ITaskService {
  getTasksByProject(projectId: number): Promise<readonly TaskItem[]>;
}

export interface IMeetingService {
  getRecentMeetings(): Promise<readonly Meeting[]>;
}

export interface IDocumentService {
  getRecentDocuments(): Promise<readonly DocumentRecord[]>;
}

export interface IPermissionService {
  getCurrentUser(): Promise<User>;
  isAdministrator(): Promise<boolean>;
}

export interface IProvisioningService {
  validateAndProvision(): Promise<ProvisioningStatus>;
}

export interface ISettingsService {
  getConfiguration(): Promise<Configuration>;
}

export interface IAppServices {
  readonly projects: IProjectService;
  readonly tasks: ITaskService;
  readonly meetings: IMeetingService;
  readonly documents: IDocumentService;
  readonly permissions: IPermissionService;
  readonly provisioning: IProvisioningService;
  readonly settings: ISettingsService;
}
