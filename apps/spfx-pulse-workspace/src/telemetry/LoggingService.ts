export class LoggingService {
  public static logError(error: Error): void {
    // Centralize logging so telemetry sinks can be added later.
    // eslint-disable-next-line no-console
    console.error('[AEWTTR]', error);
  }
}
