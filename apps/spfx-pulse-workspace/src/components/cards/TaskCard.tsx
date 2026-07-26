import * as React from 'react';
import { Card, Body1, Subtitle2 } from '@fluentui/react-components';

import { TaskItem } from '@/models/TaskItem';
import { toDisplayDate } from '@/utils/dateUtils';

interface TaskCardProps {
  readonly task: TaskItem;
}

export function TaskCard(props: TaskCardProps): React.ReactElement {
  return (
    <Card>
      <Subtitle2>{props.task.title}</Subtitle2>
      <Body1>{props.task.status}</Body1>
      <Body1>{toDisplayDate(props.task.dueDate)}</Body1>
    </Card>
  );
}
