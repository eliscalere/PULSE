import * as React from 'react';
import { Dropdown, Option } from '@fluentui/react-components';

interface ProjectFilterPanelProps {
  readonly onTeamChanged: (team: string) => void;
}

export function ProjectFilterPanel(props: ProjectFilterPanelProps): React.ReactElement {
  return (
    <Dropdown placeholder="Filter by team" onOptionSelect={(_, data) => props.onTeamChanged(String(data.optionValue ?? ''))}>
      <Option value="">All Teams</Option>
      <Option value="PM">PM</Option>
      <Option value="Engineering">Engineering</Option>
      <Option value="Cyber">Cyber</Option>
      <Option value="Support">Support</Option>
    </Dropdown>
  );
}
