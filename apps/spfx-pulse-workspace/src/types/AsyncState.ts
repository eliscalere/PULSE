export interface AsyncState<T> {
  readonly data?: T;
  readonly error?: Error;
  readonly isLoading: boolean;
}
