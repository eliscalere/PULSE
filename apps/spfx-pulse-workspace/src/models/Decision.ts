import { BaseEntity } from './BaseEntity';

export interface Decision extends BaseEntity {
  readonly projectId?: number;
  readonly decisionDate: string;
  readonly decidedById?: number;
}
