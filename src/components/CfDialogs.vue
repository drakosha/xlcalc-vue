<script setup lang="ts">
// The conditional formatting dialogs the sheet has open, each over the one it came from: the Rules Manager, a quick rule's
// dialog, New or Edit Formatting Rule, and the rule's Format Cells on top.
import { useSheetContext } from './context'
import CfManager from './CfManager.vue'
import CfQuickDialog from './CfQuickDialog.vue'
import CfRuleDialog from './CfRuleDialog.vue'
import CfStyleDialog from './CfStyleDialog.vue'

const ctx = useSheetContext()
const cf = ctx.cf
</script>

<template>
  <CfManager v-if="cf.manager.value" />
  <CfQuickDialog v-if="cf.quick.value" :key="cf.quick.value.kind" />
  <CfRuleDialog v-if="cf.rule.value" :key="cf.rule.value.base ? 'edit' : 'new'" />
  <CfStyleDialog v-if="cf.style.value" :style="cf.style.value.style" :messages="ctx.messages.value" @done="cf.styleClose($event)" @close="cf.styleClose(null)" />
</template>
