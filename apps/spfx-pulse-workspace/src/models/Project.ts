import { BaseEntity } from './BaseEntity';

export type ProjectHealth = 'Green' | 'Amber' | 'Red' | 'Blue';

export interface Project extends BaseEntity {
  readonly code: string;
  readonly description: string;
  readonly ownerId?: number;
  readonly health: ProjectHealth;
  readonly status: 'Not Started' | 'In Progress' | 'Blocked' | 'Complete';
  readonly priority: 'Immediate' | 'High' | 'Medium' | 'Low';
  readonly targetDate?: string;
  readonly team?: string;
}
