var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { getLibrarySchemas, getListSchemas, getSeedData } from '@/helpers/schemaHelper';
import { ColumnProvisioner } from './ColumnProvisioner';
import { ViewProvisioner } from './ViewProvisioner';
import { ListProvisioner } from './ListProvisioner';
import { LibraryProvisioner } from './LibraryProvisioner';
import { ValidationService } from './ValidationService';
import { UpgradeService } from './UpgradeService';
var ProvisioningService = /** @class */ (function () {
    function ProvisioningService(sp) {
        var _this = this;
        this.sp = sp;
        var columns = new ColumnProvisioner(sp, function (listKey) { return _this.resolveTitle(listKey); });
        var views = new ViewProvisioner(sp);
        this.validator = new ValidationService(sp);
        this.listProvisioner = new ListProvisioner(sp, columns, views);
        this.libraryProvisioner = new LibraryProvisioner(sp, columns, views);
        this.upgrades = new UpgradeService();
    }
    ProvisioningService.prototype.validateAndProvision = function () {
        return __awaiter(this, void 0, void 0, function () {
            var validation, _i, _a, schema, _b, _c, schema;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0: return [4 /*yield*/, this.validator.validate()];
                    case 1:
                        validation = _d.sent();
                        if (!validation.isProvisioningRequired) {
                            return [2 /*return*/, validation];
                        }
                        _i = 0, _a = getListSchemas();
                        _d.label = 2;
                    case 2:
                        if (!(_i < _a.length)) return [3 /*break*/, 5];
                        schema = _a[_i];
                        return [4 /*yield*/, this.listProvisioner.ensureList(schema)];
                    case 3:
                        _d.sent();
                        _d.label = 4;
                    case 4:
                        _i++;
                        return [3 /*break*/, 2];
                    case 5:
                        _b = 0, _c = getLibrarySchemas();
                        _d.label = 6;
                    case 6:
                        if (!(_b < _c.length)) return [3 /*break*/, 9];
                        schema = _c[_b];
                        return [4 /*yield*/, this.libraryProvisioner.ensureLibrary(schema)];
                    case 7:
                        _d.sent();
                        _d.label = 8;
                    case 8:
                        _b++;
                        return [3 /*break*/, 6];
                    case 9: return [4 /*yield*/, this.seedLists()];
                    case 10:
                        _d.sent();
                        return [4 /*yield*/, this.upgrades.runUpgradePlan()];
                    case 11:
                        _d.sent();
                        return [2 /*return*/, this.validator.validate()];
                }
            });
        });
    };
    ProvisioningService.prototype.seedLists = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _i, _a, record, listTitle, list, existing;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _i = 0, _a = getSeedData();
                        _b.label = 1;
                    case 1:
                        if (!(_i < _a.length)) return [3 /*break*/, 5];
                        record = _a[_i];
                        listTitle = this.resolveTitle(record.listKey);
                        list = this.sp.web.lists.getByTitle(listTitle);
                        return [4 /*yield*/, list.items.filter("Title eq '".concat(String(record.values.Title), "'"))()];
                    case 2:
                        existing = _b.sent();
                        if (!(existing.length === 0)) return [3 /*break*/, 4];
                        return [4 /*yield*/, list.items.add(record.values)];
                    case 3:
                        _b.sent();
                        _b.label = 4;
                    case 4:
                        _i++;
                        return [3 /*break*/, 1];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    ProvisioningService.prototype.resolveTitle = function (listKey) {
        var listSchema = getListSchemas().find(function (schema) { return schema.key === listKey; });
        var librarySchema = getLibrarySchemas().find(function (schema) { return schema.key === listKey; });
        if (listSchema) {
            return listSchema.title;
        }
        if (librarySchema) {
            return librarySchema.title;
        }
        throw new Error("Unknown schema key: ".concat(listKey));
    };
    return ProvisioningService;
}());
export { ProvisioningService };
//# sourceMappingURL=ProvisioningService.js.map