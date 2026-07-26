import * as React from 'react';

import { useAppContext } from '@/context/AppContext';
import { useAsyncData } from '@/hooks/useAsyncData';
import { LoadingOverlay } from '@/components/common/LoadingOverlay';
import { Body1, Link, Title3 } from '@fluentui/react-components';

export function DocumentsPage(): React.ReactElement {
  const { services } = useAppContext();
  const documents = useAsyncData(async () => services.documents.getRecentDocuments(), [services]);

  if (documents.isLoading) {
    return <LoadingOverlay />;
  }

  return (
    <div>
      <Title3>Documents</Title3>
      {documents.data?.map((document) => (
        <Body1 key={document.id}>
          <Link href={document.libraryUrl}>{document.title}</Link>
        </Body1>
      ))}
    </div>
  );
}
