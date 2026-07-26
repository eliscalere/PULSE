import { SPFI } from '@pnp/sp';

import { getLibrarySchemas, getListSchemas, getSeedData } from '@/helpers/schemaHelper';
import { IProvisioningService } from '@/interfaces/IAppServices';
import { ProvisioningStatus } from './ProvisioningModels';
import { ColumnProvisioner } from './ColumnProvisioner';
import { ViewProvisioner } from './ViewProvisioner';
import { ListProvisioner } from './ListProvisioner';
import { LibraryProvisioner } from './LibraryProvisioner';
import { ValidationService } from './ValidationService';
import { UpgradeService } from './UpgradeService';

export class ProvisioningService implements IProvisioningService {
  private readonly validator: ValidationService;
  private readonly listProvisioner: ListProvisioner;
  private readonly libraryProvisioner: LibraryProvisioner;
  private readonly upgrades: UpgradeService;

  public constructor(private readonly sp: SPFI) {
    const columns = new ColumnProvisioner(sp, (listKey) => this.resolveTitle(listKey));
    const views = new ViewProvisioner(sp);

    this.validator = new ValidationService(sp);
    this.listProvisioner = new ListProvisioner(sp, columns, views);
    this.libraryProvisioner = new LibraryProvisioner(sp, columns, views);
    this.upgrades = new UpgradeService();
  }

  public async validateAndProvision(): Promise<ProvisioningStatus> {
    const validation = await this.validator.validate();

    if (!validation.isProvisioningRequired) {
      return validation;
    }

    for (const schema of getListSchemas()) {
      await this.listProvisioner.ensureList(schema);
    }

    for (const schema of getLibrarySchemas()) {
      await this.libraryProvisioner.ensureLibrary(schema);
    }

    await this.seedLists();
    await this.upgrades.runUpgradePlan();

    return this.validator.validate();
  }

  private async seedLists(): Promise<void> {
    for (const record of getSeedData()) {
      const listTitle = this.resolveTitle(record.listKey);
      const list = this.sp.web.lists.getByTitle(listTitle);
      const existing = await list.items.filter(`Title eq '${String(record.values.Title)}'`)();
      if (existing.length === 0) {
        await list.items.add(record.values);
      }
    }
  }

  private resolveTitle(listKey: string): string {
    const listSchema = getListSchemas().find((schema) => schema.key === listKey);
    const librarySchema = getLibrarySchemas().find((schema) => schema.key === listKey);

    if (listSchema) {
      return listSchema.title;
    }

    if (librarySchema) {
      return librarySchema.title;
    }

    throw new Error(`Unknown schema key: ${listKey}`);
  }
}
