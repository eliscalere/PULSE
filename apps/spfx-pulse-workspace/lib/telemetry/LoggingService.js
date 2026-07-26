var LoggingService = /** @class */ (function () {
    function LoggingService() {
    }
    LoggingService.logError = function (error) {
        // Centralize logging so telemetry sinks can be added later.
        // eslint-disable-next-line no-console
        console.error('[AEWTTR]', error);
    };
    return LoggingService;
}());
export { LoggingService };
//# sourceMappingURL=LoggingService.js.map