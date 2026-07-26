import * as React from 'react';
import { Card, Body1, Subtitle2 } from '@fluentui/react-components';

import { Project } from '@/models/Project';
import { toDisplayDate } from '@/utils/dateUtils';

interface ProjectCardProps {
  readonly project: Project;
}

export function ProjectCard(props: ProjectCardProps): React.ReactElement {
  const { project } = props;
  return (
    <Card>
      <Subtitle2>{project.title}</Subtitle2>
      <Body1>{project.code}</Body1>
      <Body1>{project.health}</Body1>
      <Body1>{toDisplayDate(project.targetDate)}</Body1>
    </Card>
  );
}
