<script setup lang="ts">
defineProps<{
  fileName?: string
  sheetCount: number
  zoom: number
}>()
const emit = defineEmits<{
  (e: 'update:zoom', value: number): void
}>()

const STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2]

function setZoom(z: number) {
  emit('update:zoom', Math.min(3, Math.max(0.3, z)))
}
</script>

<template>
  <div class="toolbar">
    <span class="file" :title="fileName">{{ fileName || '未命名工作簿' }}</span>
    <span class="meta">{{ sheetCount }} 个工作表</span>
    <div class="spacer" />
    <div class="zoom">
      <button @click="setZoom(zoom - 0.1)" title="缩小">−</button>
      <select :value="zoom" @change="setZoom(parseFloat(($event.target as HTMLSelectElement).value))">
        <option v-for="s in STEPS" :key="s" :value="s">{{ Math.round(s * 100) }}%</option>
        <option v-if="!STEPS.includes(zoom)" :value="zoom">{{ Math.round(zoom * 100) }}%</option>
      </select>
      <button @click="setZoom(zoom + 0.1)" title="放大">+</button>
    </div>
  </div>
</template>

<style scoped>
.toolbar {
  display: flex;
  align-items: center;
  gap: 12px;
  height: 36px;
  padding: 0 12px;
  background: #fbfbfb;
  border-bottom: 1px solid #e2e4e7;
  font-size: 13px;
  color: #444;
  flex: 0 0 auto;
}
.file {
  font-weight: 600;
  max-width: 320px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.meta { color: #888; }
.spacer { flex: 1; }
.zoom { display: flex; align-items: center; gap: 4px; }
.zoom button {
  width: 24px;
  height: 24px;
  border: 1px solid #d0d3d6;
  background: #fff;
  border-radius: 4px;
  cursor: pointer;
  font-size: 15px;
  line-height: 1;
}
.zoom button:hover { background: #f0f2f4; }
.zoom select {
  height: 24px;
  border: 1px solid #d0d3d6;
  border-radius: 4px;
  background: #fff;
}
</style>
