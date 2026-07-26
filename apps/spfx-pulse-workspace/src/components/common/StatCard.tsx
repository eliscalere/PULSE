import * as React from 'react';
import { Card, CardHeader, Body1, Subtitle2 } from '@fluentui/react-components';

interface StatCardProps {
  readonly label: string;
  readonly value: string | number;
  readonly description: string;
}

export function StatCard(props: StatCardProps): React.ReactElement {
  return (
    <Card>
      <CardHeader header={<Subtitle2>{props.label}</Subtitle2>} />
      <Body1>{props.value}</Body1>
      <Body1>{props.description}</Body1>
    </Card>
  );
}
