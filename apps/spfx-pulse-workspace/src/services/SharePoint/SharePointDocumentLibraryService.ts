import { SPFI } from '@pnp/sp';
import '@pnp/sp/folders';
import '@pnp/sp/files/folder';

export class SharePointDocumentLibraryService {
  public constructor(private readonly sp: SPFI) {}

  public async getFiles(libraryTitle: string): Promise<readonly { Name: string; ServerRelativeUrl: string }[]> {
    return this.sp.web.lists.getByTitle(libraryTitle).rootFolder.files.select('Name', 'ServerRelativeUrl')();
  }
}
