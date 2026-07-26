import * as React from 'react';
import { Spinner } from '@fluentui/react-components';

export function LoadingOverlay(): React.ReactElement {
  return <Spinner label="Loading AEWTTR-PULSE..." />;
}
