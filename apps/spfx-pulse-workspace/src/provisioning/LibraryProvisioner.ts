import { SPFI } from '@pnp/sp';

import { SharePointLibrarySchema } from '@/types/Schema';
import { ColumnProvisioner } from './ColumnProvisioner';
import { ViewProvisioner } from './ViewProvisioner';

export class LibraryProvisioner {
  public constructor(
    private readonly sp: SPFI,
    private readonly columns: ColumnProvisioner,
    private readonly views: ViewProvisioner
  ) {}

  public async ensureLibrary(schema: SharePointLibrarySchema): Promise<void> {
    const lists = await this.sp.web.lists.select('Title')();
    const found = lists.some((list) => list.Title === schema.title);

    if (!found) {
      await this.sp.web.lists.add(schema.title, schema.description, 101);
      await this.sp.web.lists.getByTitle(schema.title).update({ EnableVersioning: schema.versioningEnabled });
    }

    await this.columns.ensureColumns(schema.title, schema.fields);
    await this.views.ensureViews(schema.title, schema.views);
  }
}
