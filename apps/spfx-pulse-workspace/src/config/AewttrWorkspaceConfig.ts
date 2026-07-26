import { Configuration } from '@/models/Configuration';

export const AewttrWorkspaceConfig: Configuration = {
  siteUrl: '',
  enableProvisioning: true,
  lists: {
    projects: 'AEWTTR Projects',
    tasks: 'AEWTTR Tasks',
    meetings: 'AEWTTR Meetings',
    risks: 'AEWTTR Risks',
    issues: 'AEWTTR Issues',
    decisions: 'AEWTTR Decisions',
    actionItems: 'AEWTTR Action Items',
    requirements: 'AEWTTR Requirements',
    milestones: 'AEWTTR Milestones',
    notifications: 'AEWTTR Notifications',
    configuration: 'AEWTTR Configuration'
  },
  libraries: {
    documents: 'AEWTTR Documents',
    reports: 'AEWTTR Reports'
  }
};
