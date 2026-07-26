import * as React from 'react';
import { Body1, Title2 } from '@fluentui/react-components';

interface PageHeaderProps {
  readonly title: string;
  readonly subtitle: string;
}

export function PageHeader(props: PageHeaderProps): React.ReactElement {
  return (
    <div>
      <Title2>{props.title}</Title2>
      <Body1>{props.subtitle}</Body1>
    </div>
  );
}
