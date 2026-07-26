export interface Configuration {
  readonly siteUrl: string;
  readonly lists: Record<string, string>;
  readonly libraries: Record<string, string>;
  readonly enableProvisioning: boolean;
}
