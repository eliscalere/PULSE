import { BaseEntity } from './BaseEntity';

export interface Issue extends BaseEntity {
  readonly projectId?: number;
  readonly status: 'Open' | 'Monitoring' | 'Resolved';
  readonly ownerId?: number;
}
