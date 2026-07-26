export function toDisplayDate(value) {
    if (!value) {
        return 'Not scheduled';
    }
    var date = new Date(value);
    return Number.isNaN(date.valueOf()) ? value : date.toLocaleDateString();
}
//# sourceMappingURL=dateUtils.js.map