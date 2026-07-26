export function toDisplayDate(value?: string): string {
  if (!value) {
    return 'Not scheduled';
  }

  const date = new Date(value);
  return Number.isNaN(date.valueOf()) ? value : date.toLocaleDateString();
}
