import * as React from 'react';
import { Table, TableBody, TableCell, TableHeader, TableHeaderCell, TableRow } from '@fluentui/react-components';

import { Project } from '@/models/Project';

interface ProjectTableProps {
  readonly projects: readonly Project[];
}

export function ProjectTable(props: ProjectTableProps): React.ReactElement {
  return (
    <Table>
      <TableHeader>
        <TableRow>
          <TableHeaderCell>Project</TableHeaderCell>
          <TableHeaderCell>Code</TableHeaderCell>
          <TableHeaderCell>Status</TableHeaderCell>
          <TableHeaderCell>Health</TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {props.projects.map((project) => (
          <TableRow key={project.id}>
            <TableCell>{project.title}</TableCell>
            <TableCell>{project.code}</TableCell>
            <TableCell>{project.status}</TableCell>
            <TableCell>{project.health}</TableCell>
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
}
