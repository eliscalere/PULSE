import { SPFI } from '@pnp/sp';

import { getLibrarySchemas, getListSchemas } from '@/helpers/schemaHelper';
import { ProvisioningMessage, ProvisioningStatus } from './ProvisioningModels';

export class ValidationService {
  public constructor(private readonly sp: SPFI) {}

  public async validate(): Promise<ProvisioningStatus> {
    const messages: ProvisioningMessage[] = [];
    const existingLists = await this.sp.web.lists.select('Title')();
    const existingTitles = new Set(existingLists.map((list) => list.Title));

    let missingCount = 0;
    for (const list of getListSchemas()) {
      if (!existingTitles.has(list.title)) {
        missingCount += 1;
        messages.push({ level: 'warning', text: `Missing list: ${list.title}` });
      }
    }

    for (const library of getLibrarySchemas()) {
      if (!existingTitles.has(library.title)) {
        missingCount += 1;
        messages.push({ level: 'warning', text: `Missing library: ${library.title}` });
      }
    }

    if (missingCount === 0) {
      messages.push({ level: 'success', text: 'All required SharePoint resources are present.' });
    }

    return {
      isReady: missingCount === 0,
      isProvisioningRequired: missingCount > 0,
      messages
    };
  }
}
