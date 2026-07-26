import * as React from 'react';
import { Body1, Button, Title3 } from '@fluentui/react-components';

import { LoggingService } from '@/telemetry/LoggingService';

interface ErrorBoundaryState {
  readonly error?: Error;
}

export class ErrorBoundary extends React.Component<React.PropsWithChildren<unknown>, ErrorBoundaryState> {
  public constructor(props: React.PropsWithChildren<unknown>) {
    super(props);
    this.state = {};
  }

  public static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }

  public componentDidCatch(error: Error): void {
    LoggingService.logError(error);
  }

  public render(): React.ReactNode {
    if (!this.state.error) {
      return this.props.children;
    }

    return (
      <div>
        <Title3>Something went wrong</Title3>
        <Body1>{this.state.error.message}</Body1>
        <Button appearance="secondary" onClick={() => window.location.reload()}>
          Reload
        </Button>
      </div>
    );
  }
}
