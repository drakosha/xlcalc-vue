// A source's error as the sheet reports it: the error's own text, and the outcome of a call that threw.
import type { EditOutcome } from './contract/source'

/** An error's message; '' for anything else, where the caller describes the failure in its own words. */
export const errorText = (e: unknown): string => (e instanceof Error ? e.message : '')

/** The outcome of a call that threw: failed, with the error's message. */
export const failedOutcome = (e: unknown): EditOutcome => ({ ok: false, reason: 'failed', message: errorText(e) })
