export function toError(error) {
    if (error instanceof Error) {
        return error;
    }
    return new Error(typeof error === 'string' ? error : 'An unexpected error occurred.');
}
//# sourceMappingURL=errorUtils.js.map