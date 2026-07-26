import { SPFI } from '@pnp/sp';

export class SharePointListService {
  public constructor(private readonly sp: SPFI) {}

  public async getItems<T>(listTitle: string, select: readonly string[]): Promise<readonly T[]> {
    return this.sp.web.lists.getByTitle(listTitle).items.select(...select)();
  }
}
