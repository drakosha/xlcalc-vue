// The answers of the host's resolver (xlcalc 0.68's `externals.resolve`) given as promises. A call of the session waits for
// them within its `timeoutMs`; past it the cells show #BUSY! or #GETTING_DATA, and each answer that arrives later tells the
// source to take it (`settleConnections`). Closing the source ends every wait, as if no answer came, and asks the host nothing more.
import type { AsyncSession, AsyncSessionOptions, EditResult, ExternalAnswer, ExternalResolver } from 'xlcalc'

const isThenable = (x: unknown): x is PromiseLike<ExternalAnswer> => x !== null && typeof x === 'object' && typeof (x as { then?: unknown }).then === 'function'

/** The functions of connections xlcalc 0.68 asks the resolver for (its `Externals`), besides the host's own `functions`. */
const CONNECTIONS = ['CUBEKPIMEMBER', 'CUBEMEMBER', 'CUBEMEMBERPROPERTY', 'CUBERANKEDMEMBER', 'CUBESET', 'CUBESETCOUNT', 'CUBEVALUE', 'RTD', 'WEBSERVICE', 'IMAGE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE', 'CALL', 'REGISTER.ID', 'EUROCONVERT']

/**
 * The deadline of a second `settleConnections` after the first ran past the calculator's: a call stopped so leaves the
 * answers in the session untaken, and another with the same deadline takes none of them (xlcalc 0.71).
 */
export const RETRY_MS = 60_000

/** The session's options with the host's resolver watched: `arrived` is called once an answer given as a promise settles. */
export function connections(options: AsyncSessionOptions | undefined, arrived: () => void) {
  let closed = false
  /** The answers on their way, each ended as none when the source closes. */
  const waits = new Set<() => void>()
  const host = options?.externals?.resolve
  const resolve: ExternalResolver | undefined =
    host &&
    ((call) => {
      // a call made while the source closes (one an answer led to) is not asked
      if (closed) return undefined
      const out = host(call)
      if (!isThenable(out)) return out
      return new Promise<ExternalAnswer>((done, fail) => {
        const none = () => done(undefined)
        waits.add(none)
        const settled = (then: () => void) => {
          if (!waits.delete(none)) return
          then()
          if (!closed) arrived()
        }
        out.then(
          (answer) => settled(() => done(answer)),
          (e: unknown) => settled(() => fail(e)),
        )
      })
    })
  /** A formula's own call outside the workbook, by its function's name as the library finds it (`name(`). */
  const names = [...new Set([...(options?.externals?.functions ?? []).map((n) => n.toUpperCase()), ...CONNECTIONS])]
  const own = new RegExp(`(?<![\\w.])(?:_xlfn\\.)?(?:${names.map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})\\(`, 'i')

  return {
    session: resolve ? { ...options, externals: { ...options!.externals!, resolve } } : options,
    /** Whether the host answers calls as the workbook calculates, so that Refresh All asks something again. */
    resolves: !!host,
    /** Whether a formula (as the session stores it) calls a function outside the workbook itself. */
    calls: (formula: string): boolean => own.test(formula),
    get closed() {
      return closed
    },
    close() {
      closed = true
      for (const none of waits) none()
      waits.clear()
    },
  }
}

/** The session's `settleConnections`; past the calculator's deadline (`soft_timeout`), once more with `RETRY_MS`. */
export async function takeAnswers(s: AsyncSession): Promise<EditResult> {
  try {
    return await s.settleConnections()
  } catch (e) {
    if ((e as { code?: string })?.code !== 'soft_timeout') throw e
  }
  return s.settleConnections({ timeoutMs: RETRY_MS })
}
