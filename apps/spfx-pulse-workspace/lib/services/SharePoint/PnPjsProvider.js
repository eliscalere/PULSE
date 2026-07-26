import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/site-users/web';
var PnPjsProvider = /** @class */ (function () {
    function PnPjsProvider() {
    }
    PnPjsProvider.create = function (context) {
        return spfi().using(SPFx(context));
    };
    return PnPjsProvider;
}());
export { PnPjsProvider };
//# sourceMappingURL=PnPjsProvider.js.map