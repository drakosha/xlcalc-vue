// One query to the source with its own signal: an exception inside it fails as a rejection does, and a response or a
// failure that arrives once the query is no longer the current one is discarded.

/** Calls `call` with a new controller's signal; `take` or `fail` run only while `live()` holds. Returns the controller, to abort it. */
export function askSource<T>(call: (signal: AbortSignal) => Promise<T>, live: () => boolean, take: (value: T) => void, fail: () => void): AbortController {
  const ctrl = new AbortController()
  const asked = async () => call(ctrl.signal)
  asked().then(
    (value) => {
      if (live()) take(value)
    },
    () => {
      if (live()) fail()
    },
  )
  return ctrl
}
