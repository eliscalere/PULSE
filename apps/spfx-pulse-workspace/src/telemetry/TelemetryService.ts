import { LoggingService } from './LoggingService';

export class TelemetryService {
  public trackException(error: Error): void {
    LoggingService.logError(error);
  }
}
