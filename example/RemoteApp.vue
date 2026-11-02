<script setup lang="ts">
// The sheet and the toolbar over a workbook a server serves (`scripts/remote-check.ts`): `?user=` goes to the server as
// a header, which decides the viewer's access; several pages are several viewers of one workbook.
import { onBeforeUnmount, shallowRef } from 'vue'
import { SheetToolbar, SheetView, ruMessages, type SheetViewApi, type SheetViewState } from '../src'
import { remoteSource, type RemoteSource } from '../src/remote'

const user = new URLSearchParams(location.search).get('user') ?? 'viewer'
const source = shallowRef<RemoteSource | null>(null)
const failed = shallowRef('')
remoteSource('/sheet/1', { headers: { 'x-user': user } }).then(
  (s) => (source.value = s),
  (e: Error) => (failed.value = e.message),
)
onBeforeUnmount(() => source.value?.close())
const state = shallowRef<SheetViewState | null>(null)
const sheet = shallowRef<SheetViewApi | null>(null)
</script>

<template>
  <div class="remote">
    <p v-if="failed" data-remote-failed>{{ failed }}</p>
    <template v-else-if="source">
      <SheetToolbar :source="source" :sheet="sheet" :state="state" :messages="ruMessages" />
      <SheetView ref="sheet" class="remote-sheet" :source="source" :messages="ruMessages" :data-remote-access="source.access" @state="state = $event" />
    </template>
  </div>
</template>

<style>
.remote {
  display: flex;
  flex-direction: column;
  height: 100%;
}
.remote-sheet {
  flex: 1;
  min-height: 0;
}
</style>
