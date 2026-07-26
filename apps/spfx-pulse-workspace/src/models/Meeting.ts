import { BaseEntity } from './BaseEntity';

export interface Meeting extends BaseEntity {
  readonly meetingDate: string;
  readonly facilitatorId?: number;
  readonly summary?: string;
  readonly nextSteps?: string;
}
