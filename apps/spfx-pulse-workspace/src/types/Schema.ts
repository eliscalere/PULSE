export type SharePointFieldType =
  | 'Text'
  | 'Note'
  | 'Number'
  | 'Currency'
  | 'DateTime'
  | 'Choice'
  | 'Lookup'
  | 'User'
  | 'Boolean';

export interface ChoiceOption {
  readonly value: string;
}

export interface LookupReference {
  readonly listKey: string;
  readonly showField: string;
}

export interface SharePointFieldSchema {
  readonly internalName: string;
  readonly displayName: string;
  readonly type: SharePointFieldType;
  readonly required?: boolean;
  readonly indexed?: boolean;
  readonly enforceUniqueValues?: boolean;
  readonly choices?: readonly ChoiceOption[];
  readonly lookup?: LookupReference;
  readonly multiValue?: boolean;
}

export interface SharePointViewSchema {
  readonly key: string;
  readonly title: string;
  readonly fields: readonly string[];
  readonly query?: string;
  readonly rowLimit?: number;
  readonly defaultView?: boolean;
}

export interface SharePointListSchema {
  readonly key: string;
  readonly title: string;
  readonly description: string;
  readonly template: number;
  readonly versioningEnabled: boolean;
  readonly fields: readonly SharePointFieldSchema[];
  readonly views: readonly SharePointViewSchema[];
}

export interface SharePointLibrarySchema extends Omit<SharePointListSchema, 'template'> {
  readonly documentTemplate?: string;
}

export interface SeedRecord {
  readonly listKey: string;
  readonly values: Record<string, string | number | boolean>;
}
