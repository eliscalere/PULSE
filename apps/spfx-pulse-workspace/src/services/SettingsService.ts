import { ISettingsService } from '@/interfaces/IAppServices';
import { Configuration } from '@/models/Configuration';
import { AewttrWorkspaceConfig } from '@/config/AewttrWorkspaceConfig';

export class SettingsService implements ISettingsService {
  public async getConfiguration(): Promise<Configuration> {
    return AewttrWorkspaceConfig;
  }
}
