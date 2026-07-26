import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/site-users/web';

export class PnPjsProvider {
  public static create(context: WebPartContext): SPFI {
    return spfi().using(SPFx(context));
  }
}
