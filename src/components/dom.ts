// Small helpers of the components' templates: an element kept in a ref, the value of the field an event came from.
import type { Ref } from 'vue'

/** A function ref that keeps the element in `target`, null once it is gone. */
export const refInto =
  <T extends Element>(target: Ref<T | null>) =>
  (el: unknown): void => {
    target.value = (el as T | null) ?? null
  }

/** The value of the field or the list an `input` or `change` event came from. */
export const fieldValue = (e: Event): string => (e.target as HTMLInputElement | HTMLSelectElement).value
