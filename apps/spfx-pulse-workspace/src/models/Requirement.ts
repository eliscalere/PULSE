import { BaseEntity } from './BaseEntity';

export interface Requirement extends BaseEntity {
  readonly projectId?: number;
  readonly source?: string;
  readonly verificationMethod?: string;
}
