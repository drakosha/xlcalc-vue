<script lang="ts">
import type { Component } from 'vue'

/** An item of a toolbar menu: its hook (`data-tb`), text, icon, whether it shows checked, a line before it, and its action. */
export interface ToolbarDropItem {
  hook: string
  label: string
  icon?: Component
  /** Excel's check on an item that shows the active cell's state; undefined for a command. */
  checked?: boolean
  line?: boolean
  run: () => void
}
</script>

<script setup lang="ts">
// A dropdown of SheetToolbar inside its `.xs-tb-split`: the arrow of a split button, or with `icon` a button that is the menu
// as a whole (Excel's Orientation), and the menu below it while `open`.
import { IconDown } from './icons'

defineProps<{
  hook: string
  /** The button's name, its ScreenTip and the menu's. */
  label: string
  open: boolean
  disabled: boolean
  items: readonly ToolbarDropItem[]
  icon?: Component
}>()
const emit = defineEmits<{ toggle: [] }>()
</script>

<template>
  <button
    type="button"
    :class="['xs-tool', icon ? 'xs-tb-wide' : 'xs-tb-caret']"
    :data-tb="hook"
    :data-tip="label"
    :aria-label="label"
    aria-haspopup="menu"
    :aria-expanded="open"
    :disabled="disabled"
    @click="emit('toggle')"
  >
    <component :is="icon" v-if="icon" />
    <IconDown :class="{ 'xs-tb-down': icon }" />
  </button>
  <div v-if="open" class="xs-tb-pop" role="menu" :aria-label="label">
    <template v-for="it in items" :key="it.hook">
      <div v-if="it.line" class="xs-tb-sep" role="separator" />
      <button
        type="button"
        class="xs-tb-item"
        :role="it.checked === undefined ? 'menuitem' : 'menuitemcheckbox'"
        :aria-checked="it.checked"
        :data-tb="it.hook"
        @click="it.run()"
      >
        <component :is="it.icon" v-if="it.icon" />
        {{ it.label }}
      </button>
    </template>
  </div>
</template>
