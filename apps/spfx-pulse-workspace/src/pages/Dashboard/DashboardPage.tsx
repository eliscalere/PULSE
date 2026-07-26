import * as React from 'react';

import { useAppContext } from '@/context/AppContext';
import { useAsyncData } from '@/hooks/useAsyncData';
import { LoadingOverlay } from '@/components/common/LoadingOverlay';
import { StatCard } from '@/components/common/StatCard';
import { ProjectCard } from '@/components/cards/ProjectCard';
import { PortfolioSummaryChart } from '@/components/charts/PortfolioSummaryChart';

export function DashboardPage(): React.ReactElement {
  const { services } = useAppContext();
  const projects = useAsyncData(async () => services.projects.getProjects(), [services]);
  const meetings = useAsyncData(async () => services.meetings.getRecentMeetings(), [services]);
  const documents = useAsyncData(async () => services.documents.getRecentDocuments(), [services]);

  if (projects.isLoading || meetings.isLoading || documents.isLoading) {
    return <LoadingOverlay />;
  }

  return (
    <div>
      <div>
        <StatCard label="Projects" value={projects.data?.length ?? 0} description="SharePoint-backed project portfolio." />
        <StatCard label="Meetings" value={meetings.data?.length ?? 0} description="Recent meeting records from the site." />
        <StatCard label="Documents" value={documents.data?.length ?? 0} description="Documents from the managed library." />
      </div>
      <PortfolioSummaryChart />
      <div>
        {projects.data?.slice(0, 3).map((project) => (
          <ProjectCard key={project.id} project={project} />
        ))}
      </div>
    </div>
  );
}
