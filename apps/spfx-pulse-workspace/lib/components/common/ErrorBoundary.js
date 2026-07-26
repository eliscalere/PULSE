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
import { Body1, Button, Title3 } from '@fluentui/react-components';
import { LoggingService } from '@/telemetry/LoggingService';
var ErrorBoundary = /** @class */ (function (_super) {
    __extends(ErrorBoundary, _super);
    function ErrorBoundary(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {};
        return _this;
    }
    ErrorBoundary.getDerivedStateFromError = function (error) {
        return { error: error };
    };
    ErrorBoundary.prototype.componentDidCatch = function (error) {
        LoggingService.logError(error);
    };
    ErrorBoundary.prototype.render = function () {
        if (!this.state.error) {
            return this.props.children;
        }
        return (React.createElement("div", null,
            React.createElement(Title3, null, "Something went wrong"),
            React.createElement(Body1, null, this.state.error.message),
            React.createElement(Button, { appearance: "secondary", onClick: function () { return window.location.reload(); } }, "Reload")));
    };
    return ErrorBoundary;
}(React.Component));
export { ErrorBoundary };
//# sourceMappingURL=ErrorBoundary.js.map