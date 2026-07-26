import * as React from 'react';

import { legacyWorkspaceHtml } from './LegacyWorkspaceHtml';

const MIN_HEIGHT: number = 1200;

export function LegacyWorkspaceFrame(): React.ReactElement {
  const iframeRef = React.useRef<HTMLIFrameElement | null>(null);
  const [height, setHeight] = React.useState<number>(MIN_HEIGHT);

  React.useEffect(() => {
    function onMessage(event: MessageEvent): void {
      if (!event.data || event.data.source !== 'aewttr-legacy-frame') {
        return;
      }

      const nextHeight = Number(event.data.height);
      if (!Number.isNaN(nextHeight) && nextHeight > 0) {
        setHeight(Math.max(MIN_HEIGHT, nextHeight + 24));
      }
    }

    window.addEventListener('message', onMessage);
    return () => window.removeEventListener('message', onMessage);
  }, []);

  return (
    <iframe
      ref={iframeRef}
      title="AEWTTR-PULSE"
      srcDoc={legacyWorkspaceHtml}
      style={{
        width: '100%',
        height: `${height}px`,
        border: '0',
        display: 'block',
        backgroundColor: '#eef0f8'
      }}
    />
  );
}
