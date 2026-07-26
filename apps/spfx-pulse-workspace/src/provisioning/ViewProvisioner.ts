import { SPFI } from '@pnp/sp';

import { SharePointViewSchema } from '@/types/Schema';

export class ViewProvisioner {
  public constructor(private readonly sp: SPFI) {}

  public async ensureViews(listTitle: string, views: readonly SharePointViewSchema[]): Promise<void> {
    const list = this.sp.web.lists.getByTitle(listTitle);
    const existingViews = await list.views.select('Title')();
    const existingTitles = new Set(existingViews.map((view) => view.Title));

    for (const view of views) {
      if (existingTitles.has(view.title)) {
        continue;
      }

      await list.views.add(view.title, false, {
        RowLimit: view.rowLimit ?? 100,
        DefaultView: view.defaultView ?? false,
        ViewQuery: view.query ?? ''
      });

      const createdView = list.views.getByTitle(view.title);
      for (const field of view.fields) {
        await createdView.fields.add(field);
      }
    }
  }
}
