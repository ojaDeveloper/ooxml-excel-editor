<script setup lang="ts">
/**
 * 操作工具栏: 渲染按钮 / 分隔线 / 下拉菜单 / 禁用态,并在宽度不足时把放不下的项折叠进「⋯ 更多」。
 * 用隐藏的 measure 行量每项宽度,配 ResizeObserver 算可见数量。
 */
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import type { ResolvedToolbarItem } from './toolbar-types'
import { TOOLBAR_ICONS, svgWrap } from './toolbar-icons'
import ToolbarMenu from './ToolbarMenu.vue'

const props = defineProps<{ items: ResolvedToolbarItem[] }>()

const containerEl = ref<HTMLElement | null>(null)
const measureEl = ref<HTMLElement | null>(null)
const widths = ref<number[]>([])
const containerW = ref(0)
const openMenuId = ref<string | null>(null)

const MORE_W = 46
const GAP = 4

function remeasure() {
  const el = measureEl.value
  if (!el) return
  widths.value = Array.from(el.children).map((c) => (c as HTMLElement).offsetWidth)
  containerW.value = containerEl.value?.clientWidth ?? 0
}

/** 能完整放下的项数(放不下则预留「更多」宽度后重算) */
const visibleCount = computed(() => {
  const cw = containerW.value
  const w = widths.value
  if (!cw || w.length !== props.items.length) return props.items.length
  let sum = 0
  let fitsAll = true
  for (let i = 0; i < props.items.length; i++) {
    sum += w[i] + GAP
    if (sum > cw) {
      fitsAll = false
      break
    }
  }
  if (fitsAll) return props.items.length
  let s = MORE_W
  let n = 0
  for (let i = 0; i < props.items.length; i++) {
    s += w[i] + GAP
    if (s > cw) break
    n++
  }
  return Math.max(0, n)
})
const visibleItems = computed(() => props.items.slice(0, visibleCount.value))
const overflowItems = computed(() => props.items.slice(visibleCount.value))

const MORE_ICON = TOOLBAR_ICONS.more

function toggleMenu(id: string) {
  openMenuId.value = openMenuId.value === id ? null : id
}
function run(it: ResolvedToolbarItem) {
  if (it.disabled) return
  it.onClick?.()
}
/** 点按钮: 有子菜单则展开,否则执行 */
function onButton(it: ResolvedToolbarItem) {
  if (it.disabled) return
  if (it.items?.length) toggleMenu(it.id)
  else {
    run(it)
    openMenuId.value = null
  }
}
function onPick(it: ResolvedToolbarItem) {
  run(it)
  openMenuId.value = null
}
function iconHtml(it: ResolvedToolbarItem): string | null {
  return it.iconSvg ? svgWrap(it.iconSvg) : null
}

let ro: ResizeObserver | null = null
function onDocPointer(e: MouseEvent) {
  if (containerEl.value && !containerEl.value.contains(e.target as Node)) openMenuId.value = null
}
onMounted(() => {
  nextTick(remeasure)
  ro = new ResizeObserver(() => (containerW.value = containerEl.value?.clientWidth ?? 0))
  if (containerEl.value) ro.observe(containerEl.value)
  document.addEventListener('mousedown', onDocPointer)
})
onBeforeUnmount(() => {
  ro?.disconnect()
  document.removeEventListener('mousedown', onDocPointer)
})
// 项变化(标签/数量) → 重新量宽
watch(
  () => props.items.map((i) => i.id + (i.label ?? '')).join('|'),
  () => nextTick(remeasure),
)
</script>

<template>
  <div class="action-toolbar" ref="containerEl">
    <!-- 隐藏测量行: 用于量每项宽度 -->
    <div class="measure" ref="measureEl" aria-hidden="true">
      <template v-for="it in items" :key="'m' + it.id">
        <span v-if="it.type === 'separator'" class="divider" />
        <button v-else class="tool">
          <span v-if="iconHtml(it)" class="ic" v-html="iconHtml(it)" />
          <span v-else-if="it.icon" class="ic-e">{{ it.icon }}</span>
          <span v-if="it.label" class="lb">{{ it.label }}</span>
          <span v-if="it.items" class="caret" v-html="svgWrap(TOOLBAR_ICONS.caret)" />
        </button>
      </template>
    </div>

    <!-- 可见项 -->
    <template v-for="it in visibleItems" :key="it.id">
      <span v-if="it.type === 'separator'" class="divider" />
      <div v-else class="dd">
        <button
          class="tool"
          :class="{ active: it.active, open: openMenuId === it.id }"
          :disabled="it.disabled"
          :title="it.title || it.label || it.id"
          @click="onButton(it)"
        >
          <span v-if="iconHtml(it)" class="ic" v-html="iconHtml(it)" />
          <span v-else-if="it.icon" class="ic-e">{{ it.icon }}</span>
          <span v-if="it.label" class="lb">{{ it.label }}</span>
          <span v-if="it.items" class="caret" v-html="svgWrap(TOOLBAR_ICONS.caret)" />
        </button>
        <ToolbarMenu v-if="it.items && openMenuId === it.id" :items="it.items" @pick="onPick" />
      </div>
    </template>

    <!-- 溢出折叠 -->
    <div v-if="overflowItems.length" class="dd more">
      <button
        class="tool"
        :class="{ open: openMenuId === '__more' }"
        title="更多"
        @click="toggleMenu('__more')"
      >
        <span class="ic" v-html="svgWrap(MORE_ICON)" />
      </button>
      <ToolbarMenu v-if="openMenuId === '__more'" :items="overflowItems" @pick="onPick" />
    </div>
  </div>
</template>

<style scoped>
.action-toolbar {
  position: relative;
  display: flex;
  align-items: center;
  gap: 4px;
  height: 34px;
  flex: 0 0 auto;
  padding: 0 10px;
  background: #fff;
  border-bottom: 1px solid #e2e4e7;
  overflow: visible;
}
.measure {
  position: absolute;
  left: 0;
  top: 0;
  visibility: hidden;
  pointer-events: none;
  height: 0;
  overflow: hidden;
  display: flex;
  gap: 4px;
}
.dd { position: relative; display: inline-flex; }
.tool {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  height: 26px;
  padding: 0 9px;
  border: 1px solid transparent;
  border-radius: 5px;
  background: none;
  cursor: pointer;
  font-size: 13px;
  color: #444;
  white-space: nowrap;
}
.tool:hover:not(:disabled) { background: #f0f2f4; }
.tool:disabled { color: #bbb; cursor: default; }
.tool.active { background: #e8f0fe; color: #1a73e8; border-color: #c5d9fb; }
.tool.open { background: #eef0f2; }
.tool .ic { display: inline-flex; }
.tool .ic-e { font-size: 13px; line-height: 1; }
.tool .caret { display: inline-flex; opacity: 0.6; margin-left: -2px; }
.tool .caret :deep(svg) { width: 12px; height: 12px; }
.divider {
  width: 1px;
  align-self: stretch;
  margin: 6px 4px;
  background: #e2e4e7;
}
</style>
