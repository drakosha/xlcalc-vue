// The answers of the host's resolver (xlcalc 0.68's `externals.resolve`) given as promises. A call of the session waits for
// them within its `timeoutMs`; past it the cells show #BUSY! or #GETTING_DATA, and each answer that arrives later tells the
// source to take it (`settleConnections`). Closing the source ends every wait, as if no answer came.
import type { AsyncSessionOptions, ExternalAnswer, ExternalResolver } from 'xlcalc'

const isThenable = (x: unknown): x is PromiseLike<ExternalAnswer> => x !== null && typeof x === 'object' && typeof (x as { then?: unknown }).then === 'function'

/** The session's options with the host's resolver watched: `arrived` is called once an answer given as a promise settles. */
export function connections(options: AsyncSessionOptions | undefined, arrived: () => void) {
  let closed = false
  /** The answers on their way, each ended as none when the source closes. */
  const waits = new Set<() => void>()
  const host = options?.externals?.resolve
  const resolve: ExternalResolver | undefined =
    host &&
    ((call) => {
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

  return {
    session: resolve ? { ...options, externals: { ...options!.externals!, resolve } } : options,
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
