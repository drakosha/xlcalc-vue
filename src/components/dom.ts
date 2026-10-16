// Small helpers of the components: an element kept in a ref, the value of the field an event came from, a host's listener.
import type { ComponentInternalInstance, Ref } from 'vue'

/** A function ref that keeps the element in `target`, null once it is gone. */
export const refInto =
  <T extends Element>(target: Ref<T | null>) =>
  (el: unknown): void => {
    target.value = (el as T | null) ?? null
  }

/** The value of the field or the list an `input` or `change` event came from. */
export const fieldValue = (e: Event): string => (e.target as HTMLInputElement | HTMLSelectElement).value

/** Whether the host listens to an event of the component now (`handler` as on its vnode: `onOpenLink`). */
export function hasListener(instance: ComponentInternalInstance | null, handler: string): boolean {
  const listener = instance?.vnode.props?.[handler]
  return typeof listener === 'function' || Array.isArray(listener)
}
