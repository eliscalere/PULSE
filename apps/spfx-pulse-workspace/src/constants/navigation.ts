import { AppRoute } from '@/types/AppRoute';

export interface NavigationItem {
  readonly key: AppRoute;
  readonly label: string;
}

export const NAVIGATION_ITEMS: readonly NavigationItem[] = [
  { key: 'dashboard', label: 'Dashboard' },
  { key: 'projects', label: 'Projects' },
  { key: 'meetings', label: 'Meetings' },
  { key: 'calendar', label: 'Calendar' },
  { key: 'tasks', label: 'Tasks' },
  { key: 'reports', label: 'Reports' },
  { key: 'documents', label: 'Documents' },
  { key: 'settings', label: 'Settings' },
  { key: 'admin', label: 'Admin' }
];
