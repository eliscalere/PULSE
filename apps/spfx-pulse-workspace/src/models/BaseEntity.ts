export interface BaseEntity {
  readonly id: number;
  readonly title: string;
  readonly created?: string;
  readonly modified?: string;
}
