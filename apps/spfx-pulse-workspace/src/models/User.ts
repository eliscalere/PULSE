export interface User {
  readonly id: number;
  readonly title: string;
  readonly email: string;
  readonly role: string;
  readonly isSiteOwner: boolean;
  readonly isSiteMember: boolean;
  readonly isAdministrator: boolean;
}
