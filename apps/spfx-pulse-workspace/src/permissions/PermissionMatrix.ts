export const PermissionMatrix = {
  administrator: ['read', 'write', 'configure', 'provision'],
  member: ['read', 'write'],
  reader: ['read']
} as const;
