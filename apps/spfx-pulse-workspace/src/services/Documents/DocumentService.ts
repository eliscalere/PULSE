import { SPFI } from '@pnp/sp';

import { LIBRARY_KEYS } from '@/constants/sharePointResources';
import { IDocumentService } from '@/interfaces/IAppServices';
import { DocumentRecord } from '@/models/DocumentRecord';
import { SharePointDocumentLibraryService } from '@/services/SharePoint/SharePointDocumentLibraryService';
import { AewttrWorkspaceConfig } from '@/config/AewttrWorkspaceConfig';

export class DocumentService implements IDocumentService {
  private readonly libraries: SharePointDocumentLibraryService;

  public constructor(sp: SPFI) {
    this.libraries = new SharePointDocumentLibraryService(sp);
  }

  public async getRecentDocuments(): Promise<readonly DocumentRecord[]> {
    const files = await this.libraries.getFiles(AewttrWorkspaceConfig.libraries[LIBRARY_KEYS.documents]);

    return files.slice(0, 10).map((file, index) => ({
      id: index + 1,
      title: file.Name,
      status: 'Draft',
      libraryUrl: file.ServerRelativeUrl
    }));
  }
}
