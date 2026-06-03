<script setup lang="ts">
import type { WorkbookModel } from '@/core/model/types'
import { computed } from 'vue'

const props = defineProps<{
  workbook: WorkbookModel
  active: number
}>()
const emit = defineEmits<{ (e: 'select', index: number): void }>()

// 只展示可见 sheet
const visibleSheets = computed(() => props.workbook.sheets.filter((s) => s.state === 'visible'))
</script>

<template>
  <div class="sheet-tabs">
    <button
      v-for="sheet in visibleSheets"
      :key="sheet.index"
      class="tab"
      :class="{ active: sheet.index === active }"
      @click="emit('select', sheet.index)"
      :title="sheet.name"
    >
      {{ sheet.name }}
    </button>
  </div>
</template>

<style scoped>
.sheet-tabs {
  display: flex;
  align-items: stretch;
  height: 30px;
  background: #f3f4f6;
  border-top: 1px solid #d8dadd;
  overflow-x: auto;
  flex: 0 0 auto;
}
.tab {
  border: none;
  background: transparent;
  border-right: 1px solid #e0e2e5;
  padding: 0 16px;
  font-size: 12.5px;
  color: #555;
  cursor: pointer;
  white-space: nowrap;
  max-width: 200px;
  overflow: hidden;
  text-overflow: ellipsis;
}
.tab:hover { background: #e9ebee; }
.tab.active {
  background: #fff;
  color: #166534;
  font-weight: 600;
  box-shadow: inset 0 2px 0 #21a366;
}
</style>
