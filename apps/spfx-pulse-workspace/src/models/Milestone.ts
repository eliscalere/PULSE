import { BaseEntity } from './BaseEntity';

export interface Milestone extends BaseEntity {
  readonly projectId?: number;
  readonly dueDate?: string;
  readonly status: 'Upcoming' | 'On Track' | 'At Risk' | 'Complete';
}
