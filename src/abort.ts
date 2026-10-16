// A request stopped by its signal, as `fetch` stops one: the error the sources throw and the sheet tells from a failure.

/** The error of a request its signal stopped: a DOMException named AbortError. */
export const abortError = (): DOMException => new DOMException('The request was aborted', 'AbortError')

/** Whether an error is that of a stopped request, which the sheet drops without a message. */
export const isAbort = (e: unknown): boolean => e instanceof DOMException && e.name === 'AbortError'

/** Throws the abort error once the signal is aborted. */
export function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw abortError()
}
