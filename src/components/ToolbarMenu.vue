<script setup lang="ts">
/** 工具栏下拉菜单(下拉项 / 溢出"更多"共用)。 */
import type { ResolvedToolbarItem } from './toolbar-types'
import { svgWrap } from './toolbar-icons'

defineProps<{ items: ResolvedToolbarItem[] }>()
const emit = defineEmits<{ (e: 'pick', it: ResolvedToolbarItem): void }>()
</script>

<template>
  <div class="tb-menu">
    <template v-for="it in items" :key="it.id">
      <div v-if="it.type === 'separator'" class="sep" />
      <button
        v-else
        class="mi"
        :class="{ active: it.active }"
        :disabled="it.disabled"
        @click="emit('pick', it)"
      >
        <span v-if="it.iconSvg" class="ic" v-html="svgWrap(it.iconSvg)" />
        <span v-else-if="it.icon" class="ic-e">{{ it.icon }}</span>
        <span class="lb">{{ it.label || it.id }}</span>
      </button>
    </template>
  </div>
</template>

<style scoped>
.tb-menu {
  position: absolute;
  top: 30px;
  left: 0;
  z-index: 20;
  min-width: 160px;
  background: #fff;
  border: 1px solid #d8dbde;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.14);
  padding: 4px;
  display: flex;
  flex-direction: column;
}
.mi {
  display: flex;
  align-items: center;
  gap: 8px;
  text-align: left;
  padding: 7px 10px;
  border: none;
  background: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  color: #333;
  white-space: nowrap;
}
.mi:hover:not(:disabled) { background: #f0f4ff; }
.mi:disabled { color: #bbb; cursor: default; }
.mi.active { color: #1a73e8; }
.mi .ic { display: inline-flex; width: 16px; }
.mi .ic-e { width: 16px; text-align: center; }
.sep { height: 1px; background: #eef0f2; margin: 4px 2px; }
</style>
