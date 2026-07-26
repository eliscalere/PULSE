import { BaseEntity } from './BaseEntity';

export interface ActionItem extends BaseEntity {
  readonly meetingId?: number;
  readonly ownerId?: number;
  readonly dueDate?: string;
  readonly status: 'Open' | 'In Progress' | 'Closed';
}
