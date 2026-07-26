import { BaseEntity } from './BaseEntity';

export interface TaskItem extends BaseEntity {
  readonly projectId?: number;
  readonly assigneeId?: number;
  readonly dueDate?: string;
  readonly status: 'To Do' | 'In Progress' | 'Blocked' | 'Done';
  readonly percentComplete: number;
  readonly priority: 'Immediate' | 'High' | 'Medium' | 'Low';
}
