import { SPFI } from '@pnp/sp';

export class SearchService {
  public constructor(private readonly sp: SPFI) {}

  public async search(queryText: string): Promise<readonly { title: string; url: string }[]> {
    const siteUrl = await this.sp.web.select('Url')();
    return queryText
      ? [{ title: `Search placeholder for "${queryText}"`, url: String(siteUrl.Url) }]
      : [];
  }
}
