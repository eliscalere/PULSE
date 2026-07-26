import { SharePointListSchema, SharePointLibrarySchema, SeedRecord } from '@/types/Schema';

import listsSchema from '@/schema/lists.json';
import librariesSchema from '@/schema/libraries.json';
import seedDataSchema from '@/schema/seedData.json';

export function getListSchemas(): readonly SharePointListSchema[] {
  return listsSchema as unknown as readonly SharePointListSchema[];
}

export function getLibrarySchemas(): readonly SharePointLibrarySchema[] {
  return librariesSchema as unknown as readonly SharePointLibrarySchema[];
}

export function getSeedData(): readonly SeedRecord[] {
  return seedDataSchema as unknown as readonly SeedRecord[];
}
