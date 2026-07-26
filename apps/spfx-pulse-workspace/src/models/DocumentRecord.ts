import { BaseEntity } from './BaseEntity';

export interface DocumentRecord extends BaseEntity {
  readonly projectId?: number;
  readonly libraryUrl?: string;
  readonly status: 'Draft' | 'Review' | 'Concurrence' | 'Final';
  readonly submittedById?: number;
}
