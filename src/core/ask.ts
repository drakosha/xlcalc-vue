// One question to the source with a signal of its own: a throw inside it fails as a rejection does, and an answer or a
// failure that comes once the question is no longer the one in force is dropped.

/** Ask `call` with a new controller's signal; `take` or `fail` run only while `live()` holds. Gives the controller, to abort it. */
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
