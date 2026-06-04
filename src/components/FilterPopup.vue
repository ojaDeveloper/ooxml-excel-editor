<script setup lang="ts">
/** 自动筛选浮层: 搜值 + (全选) + 去重值复选 + 确定/清除/取消。 */
import { computed, onBeforeUnmount, onMounted, reactive, ref } from 'vue'

const props = defineProps<{
  values: string[] // 该列去重值(已排序)
  selected: string[] // 当前勾选(允许)的值;空数组视为全选
  x: number
  y: number
  sortDir?: 'asc' | 'desc' | null // 该列当前排序方向(高亮用)
}>()
const emit = defineEmits<{
  (e: 'apply', checked: string[]): void
  (e: 'clear'): void
  (e: 'close'): void
  (e: 'sort', dir: 'asc' | 'desc'): void
}>()

const search = ref('')
// 本地勾选集
const checked = reactive(new Set<string>(props.selected.length ? props.selected : props.values))

const shown = computed(() => {
  const q = search.value.trim().toLowerCase()
  return q ? props.values.filter((v) => v.toLowerCase().includes(q)) : props.values
})
const allShownChecked = computed(() => shown.value.length > 0 && shown.value.every((v) => checked.has(v)))

function toggle(v: string) {
  checked.has(v) ? checked.delete(v) : checked.add(v)
}
function toggleAll() {
  const target = !allShownChecked.value
  for (const v of shown.value) target ? checked.add(v) : checked.delete(v)
}
function onOk() {
  emit('apply', [...checked])
}

const rootEl = ref<HTMLElement | null>(null)
function onDocPointer(e: MouseEvent) {
  if (rootEl.value && !rootEl.value.contains(e.target as Node)) emit('close')
}
onMounted(() => setTimeout(() => document.addEventListener('mousedown', onDocPointer), 0))
onBeforeUnmount(() => document.removeEventListener('mousedown', onDocPointer))
</script>

<template>
  <div class="filter-pop" ref="rootEl" :style="{ left: x + 'px', top: y + 'px' }" @keydown.stop>
    <div class="sort">
      <button :class="{ on: sortDir === 'asc' }" @click="emit('sort', 'asc')" title="升序">↑ 升序</button>
      <button :class="{ on: sortDir === 'desc' }" @click="emit('sort', 'desc')" title="降序">↓ 降序</button>
    </div>
    <input class="search" type="text" placeholder="搜索…" v-model="search" />
    <label class="all">
      <input type="checkbox" :checked="allShownChecked" @change="toggleAll" />
      (全选)
    </label>
    <div class="list">
      <label v-for="v in shown" :key="v" class="row">
        <input type="checkbox" :checked="checked.has(v)" @change="toggle(v)" />
        <span class="val" :title="v">{{ v }}</span>
      </label>
      <div v-if="!shown.length" class="empty">无匹配值</div>
    </div>
    <div class="foot">
      <button class="link" @click="emit('clear')">清除筛选</button>
      <div class="grow" />
      <button @click="emit('close')">取消</button>
      <button class="primary" @click="onOk">确定</button>
    </div>
  </div>
</template>

<style scoped>
.filter-pop {
  position: absolute;
  z-index: 14;
  width: 220px;
  background: #fff;
  border: 1px solid #cfd3d7;
  border-radius: 7px;
  box-shadow: 0 6px 22px rgba(0, 0, 0, 0.2);
  padding: 8px;
  font-size: 13px;
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.sort {
  display: flex;
  gap: 6px;
}
.sort button {
  flex: 1;
  height: 26px;
  border: 1px solid #d0d3d6;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  font-size: 12px;
  color: #444;
}
.sort button:hover { background: #f0f2f4; }
.sort button.on { background: #d6e9d9; border-color: #21a366; color: #1b7a4d; font-weight: 600; }
.search {
  height: 26px;
  border: 1px solid #d0d3d6;
  border-radius: 4px;
  padding: 0 8px;
  outline: none;
}
.search:focus { border-color: #1a73e8; }
.all {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 2px 2px;
  color: #444;
  font-weight: 600;
  cursor: pointer;
}
.list {
  max-height: 220px;
  overflow: auto;
  border: 1px solid #eceef0;
  border-radius: 4px;
  padding: 2px;
}
.row {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 3px 4px;
  border-radius: 3px;
  cursor: pointer;
}
.row:hover { background: #f3f6fb; }
.val {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.empty { color: #aaa; text-align: center; padding: 10px; }
.foot {
  display: flex;
  align-items: center;
  gap: 6px;
}
.foot .grow { flex: 1; }
.foot button {
  height: 27px;
  padding: 0 12px;
  border: 1px solid #d0d3d6;
  border-radius: 5px;
  background: #fff;
  cursor: pointer;
  font-size: 13px;
}
.foot button:hover { background: #f0f2f4; }
.foot button.primary { background: #21a366; border-color: #21a366; color: #fff; }
.foot button.primary:hover { background: #1c8f59; }
.foot button.link { border-color: transparent; color: #1a73e8; padding: 0 4px; }
</style>
