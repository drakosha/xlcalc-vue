// A source's error as the sheet tells it: the error's own text, and the outcome of a call that threw.
import type { EditOutcome } from './source'

/** An error's message; '' for anything else, where the caller says what failed in its own words. */
export const errorText = (e: unknown): string => (e instanceof Error ? e.message : '')

/** The outcome of a call that threw: failed, with the error's message. */
export const failedOutcome = (e: unknown): EditOutcome => ({ ok: false, reason: 'failed', message: errorText(e) })
