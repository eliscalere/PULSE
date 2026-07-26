import { LoggingService } from './LoggingService';
var TelemetryService = /** @class */ (function () {
    function TelemetryService() {
    }
    TelemetryService.prototype.trackException = function (error) {
        LoggingService.logError(error);
    };
    return TelemetryService;
}());
export { TelemetryService };
//# sourceMappingURL=TelemetryService.js.map