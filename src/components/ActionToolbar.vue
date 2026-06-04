<script setup lang="ts">
/** 操作工具栏: 渲染解析后的按钮列表(查找/筛选/排序 + 插件/自定义项),内置与外来项间加分隔线。 */
import type { ResolvedToolbarItem } from './toolbar-types'

defineProps<{ items: ResolvedToolbarItem[] }>()
</script>

<template>
  <div class="action-toolbar">
    <template v-for="(it, i) in items" :key="it.id">
      <span v-if="i > 0 && items[i - 1].kind === 'builtin' && it.kind !== 'builtin'" class="divider" />
      <button class="tool" :class="{ active: it.active }" :title="it.title || it.label || it.id" @click="it.onClick()">
        <span v-if="it.icon" class="ic">{{ it.icon }}</span>
        <span v-if="it.label" class="lb">{{ it.label }}</span>
      </button>
    </template>
  </div>
</template>

<style scoped>
.action-toolbar {
  display: flex;
  align-items: center;
  gap: 4px;
  height: 34px;
  flex: 0 0 auto;
  padding: 0 10px;
  background: #fff;
  border-bottom: 1px solid #e2e4e7;
}
.tool {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  height: 26px;
  padding: 0 10px;
  border: 1px solid transparent;
  border-radius: 5px;
  background: none;
  cursor: pointer;
  font-size: 13px;
  color: #444;
}
.tool:hover { background: #f0f2f4; }
.tool.active { background: #e8f0fe; color: #1a73e8; border-color: #c5d9fb; }
.tool .ic { font-size: 13px; line-height: 1; }
.divider {
  width: 1px;
  align-self: stretch;
  margin: 6px 4px;
  background: #e2e4e7;
}
</style>
