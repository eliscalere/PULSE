import { BaseEntity } from './BaseEntity';

export interface Risk extends BaseEntity {
  readonly projectId?: number;
  readonly probability: 'Low' | 'Medium' | 'High';
  readonly impact: 'Low' | 'Medium' | 'High';
  readonly mitigation?: string;
}
