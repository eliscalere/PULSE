export interface Notification {
  readonly id: string;
  readonly title: string;
  readonly intent: 'success' | 'info' | 'warning' | 'error';
  readonly message: string;
}
