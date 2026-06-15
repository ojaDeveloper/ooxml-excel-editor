<script setup lang="ts">
/** 查找条: 输入 + 命中计数 + 上/下一个 + 区分大小写/全字匹配 + 关闭。受控组件。 */
import { onMounted, ref } from 'vue'

defineProps<{
  query: string
  matchCount: number
  current: number // 0-based;无命中为 -1
  matchCase: boolean
  wholeCell: boolean
  editable: boolean // 开 editable 才显示替换行(1.11.0)
  replace: string
}>()
const emit = defineEmits<{
  (e: 'update:query', v: string): void
  (e: 'update:matchCase', v: boolean): void
  (e: 'update:wholeCell', v: boolean): void
  (e: 'update:replace', v: string): void
  (e: 'next'): void
  (e: 'prev'): void
  (e: 'replace-one'): void
  (e: 'replace-all'): void
  (e: 'close'): void
}>()

const inputEl = ref<HTMLInputElement | null>(null)
onMounted(() => inputEl.value?.focus())

function onKeydown(e: KeyboardEvent) {
  if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
    e.preventDefault() // 已在查找,别触发浏览器原生查找
    inputEl.value?.select()
  } else if (e.key === 'Enter') {
    e.preventDefault()
    e.shiftKey ? emit('prev') : emit('next')
  } else if (e.key === 'Escape') {
    e.preventDefault()
    emit('close')
  }
}
function onReplaceKeydown(e: KeyboardEvent) {
  if (e.key === 'Enter') { e.preventDefault(); emit('replace-one') }
  else if (e.key === 'Escape') { e.preventDefault(); emit('close') }
}
defineExpose({ focus: () => inputEl.value?.focus() })
</script>

<template>
  <div class="find-bar" @keydown.stop>
    <div class="row">
      <input
        ref="inputEl"
        class="q"
        type="text"
        placeholder="查找…"
        :value="query"
        @input="emit('update:query', ($event.target as HTMLInputElement).value)"
        @keydown="onKeydown"
      />
      <span class="count" :class="{ none: query && matchCount === 0 }">
        {{ matchCount ? `${current + 1}/${matchCount}` : query ? '无结果' : '' }}
      </span>
      <button
        class="opt"
        :class="{ on: matchCase }"
        title="区分大小写"
        @click="emit('update:matchCase', !matchCase)"
      >Aa</button>
      <button
        class="opt"
        :class="{ on: wholeCell }"
        title="全字匹配(整格相等)"
        @click="emit('update:wholeCell', !wholeCell)"
      >▢</button>
      <button class="nav" title="上一个 (Shift+Enter)" :disabled="!matchCount" @click="emit('prev')">▲</button>
      <button class="nav" title="下一个 (Enter)" :disabled="!matchCount" @click="emit('next')">▼</button>
      <button class="close" title="关闭 (Esc)" @click="emit('close')">×</button>
    </div>
    <div v-if="editable" class="row replace-row">
      <input
        class="q"
        type="text"
        placeholder="替换为…"
        :value="replace"
        @input="emit('update:replace', ($event.target as HTMLInputElement).value)"
        @keydown="onReplaceKeydown"
      />
      <button class="rep" title="替换当前并查找下一个" :disabled="!matchCount" @click="emit('replace-one')">替换</button>
      <button class="rep" title="全部替换" :disabled="!matchCount" @click="emit('replace-all')">全部替换</button>
    </div>
  </div>
</template>

<style scoped>
.find-bar {
  position: absolute;
  top: 8px;
  right: 16px;
  z-index: 12;
  display: flex;
  flex-direction: column;
  gap: 4px;
  padding: 5px 7px;
  background: #fff;
  border: 1px solid #d4d7da;
  border-radius: 7px;
  box-shadow: 0 3px 12px rgba(0, 0, 0, 0.16);
  font-size: 13px;
}
.find-bar .row {
  display: flex;
  align-items: center;
  gap: 4px;
}
.rep {
  height: 24px;
  padding: 0 8px;
  border: 1px solid #d0d3d6;
  background: #f7f8fa;
  border-radius: 4px;
  cursor: pointer;
  color: #444;
  font-size: 12px;
}
.rep:hover:not(:disabled) { background: #eef0f2; }
.rep:disabled { color: #bbb; cursor: default; }
.q {
  width: 160px;
  height: 26px;
  border: 1px solid #d0d3d6;
  border-radius: 4px;
  padding: 0 8px;
  outline: none;
}
.q:focus { border-color: #1a73e8; }
.count {
  min-width: 46px;
  text-align: center;
  color: #888;
  font-variant-numeric: tabular-nums;
}
.count.none { color: #d23; }
.opt, .nav, .close {
  height: 24px;
  min-width: 24px;
  border: 1px solid transparent;
  background: none;
  border-radius: 4px;
  cursor: pointer;
  color: #555;
  font-size: 13px;
  line-height: 1;
}
.opt.on { background: #e8f0fe; color: #1a73e8; border-color: #c5d9fb; }
.opt:hover, .nav:hover:not(:disabled), .close:hover { background: #f0f2f4; }
.nav:disabled { color: #ccc; cursor: default; }
.close { font-size: 18px; color: #999; }
</style>
