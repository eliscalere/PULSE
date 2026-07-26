import * as React from 'react';

import { useAppContext } from '@/context/AppContext';
import { useAsyncData } from '@/hooks/useAsyncData';
import { LoadingOverlay } from '@/components/common/LoadingOverlay';
import { ProjectFilterPanel } from '@/components/forms/ProjectFilterPanel';
import { ProjectTable } from '@/components/tables/ProjectTable';

export function ProjectsPage(): React.ReactElement {
  const { services } = useAppContext();
  const [teamFilter, setTeamFilter] = React.useState<string>('');
  const projects = useAsyncData(async () => services.projects.getProjects(), [services]);

  if (projects.isLoading) {
    return <LoadingOverlay />;
  }

  const filtered = (projects.data ?? []).filter((project) => !teamFilter || project.team === teamFilter);

  return (
    <div>
      <ProjectFilterPanel onTeamChanged={setTeamFilter} />
      <ProjectTable projects={filtered} />
    </div>
  );
}
