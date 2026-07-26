import { useAsyncData } from './useAsyncData';
import { useAppContext } from '@/context/AppContext';

export function usePermissions() {
  const { services } = useAppContext();
  return useAsyncData(async () => services.permissions.getCurrentUser(), [services]);
}
