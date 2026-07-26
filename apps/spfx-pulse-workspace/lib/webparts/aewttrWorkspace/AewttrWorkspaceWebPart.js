var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { AewttrWorkspaceApp } from './components/AewttrWorkspaceApp';
var AewttrWorkspaceWebPart = /** @class */ (function (_super) {
    __extends(AewttrWorkspaceWebPart, _super);
    function AewttrWorkspaceWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        return _this;
    }
    AewttrWorkspaceWebPart.prototype.render = function () {
        var element = React.createElement(AewttrWorkspaceApp, {
            context: this.context,
            isDarkTheme: this._isDarkTheme,
            description: this.properties.description
        });
        ReactDom.render(element, this.domElement);
    };
    AewttrWorkspaceWebPart.prototype.onThemeChanged = function (currentTheme) {
        this._isDarkTheme = Boolean(currentTheme === null || currentTheme === void 0 ? void 0 : currentTheme.isInverted);
    };
    AewttrWorkspaceWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    return AewttrWorkspaceWebPart;
}(BaseClientSideWebPart));
export default AewttrWorkspaceWebPart;
//# sourceMappingURL=AewttrWorkspaceWebPart.js.map