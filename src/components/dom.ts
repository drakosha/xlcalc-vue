// Small component helpers: an element kept in a ref, the value of the input an event came from, a host's listener.
import type { ComponentInternalInstance, Ref } from 'vue'

/** A function ref that keeps the element in `target`, null once it is removed. */
export const refInto =
  <T extends Element>(target: Ref<T | null>) =>
  (el: unknown): void => {
    target.value = (el as T | null) ?? null
  }

/** The value of the input or list an `input` or `change` event came from. */
export const fieldValue = (e: Event): string => (e.target as HTMLInputElement | HTMLSelectElement).value

/** Whether the host currently listens to an event of the component (`handler` as on its vnode: `onOpenLink`). */
export function hasListener(instance: ComponentInternalInstance | null, handler: string): boolean {
  const listener = instance?.vnode.props?.[handler]
  return typeof listener === 'function' || Array.isArray(listener)
}
