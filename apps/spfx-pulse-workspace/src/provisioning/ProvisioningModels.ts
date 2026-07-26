export interface ProvisioningMessage {
  readonly level: 'info' | 'warning' | 'error' | 'success';
  readonly text: string;
}

export interface ProvisioningStatus {
  readonly isReady: boolean;
  readonly isProvisioningRequired: boolean;
  readonly messages: readonly ProvisioningMessage[];
}
