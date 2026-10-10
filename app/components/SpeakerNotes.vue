<script setup lang="ts">
import type { CSSProperties } from 'vue'
import type { PptxHyperlink, PptxParagraph, PptxRun } from '~/utils/pptx-model'
import { paragraphEntries } from '~/utils/paragraph-bullets'
import { cssFontFamily, scriptSegments, type ScriptFont } from '~/utils/text-fonts'

defineProps<{ paragraphs: PptxParagraph[] }>()
const emit = defineEmits<{ 'activate-hyperlink': [link: PptxHyperlink] }>()

const spacing = (value?: { unit: 'emu' | 'ratio'; value: number }) => value
  ? value.unit === 'emu' ? `${value.value / 12700}pt` : `${value.value}em`
  : undefined

function paragraphStyle(paragraph: PptxParagraph): CSSProperties {
  return {
    fontSize: paragraph.runs[0]?.fontSizePt === undefined ? undefined : `${paragraph.runs[0].fontSizePt}pt`,
    color: paragraph.runs[0]?.color,
    textAlign: paragraph.align,
    direction: paragraph.direction,
    marginLeft: paragraph.marginLeft === undefined ? undefined : `${paragraph.marginLeft / 12700}pt`,
    marginRight: paragraph.marginRight === undefined ? undefined : `${paragraph.marginRight / 12700}pt`,
    textIndent: paragraph.indent === undefined ? undefined : `${paragraph.indent / 12700}pt`,
    marginTop: spacing(paragraph.spaceBefore),
    marginBottom: spacing(paragraph.spaceAfter),
    lineHeight: spacing(paragraph.lineSpacing),
  }
}

function bulletStyle(paragraph: PptxParagraph): CSSProperties {
  const size = paragraph.bullet?.size
  return { fontSize: size ? size.unit === 'ratio' ? `${size.value * 100}%` : `${size.value}pt` : undefined }
}

function segmentFont(run: PptxRun, script: ScriptFont): string {
  return cssFontFamily((script === 'eastAsia' ? run.fontFamilyEastAsia : script === 'complexScript' ? run.fontFamilyComplexScript : run.fontFamily) || run.fontFamily)
}

function runStyle(run: PptxRun): CSSProperties {
  const angle = (run.shadow?.angle || 0) * Math.PI / 180
  const distance = (run.shadow?.distance || 0) / 12700
  return {
    color: run.color,
    fontFamily: cssFontFamily(run.fontFamily),
    fontSize: run.fontSizePt === undefined ? undefined : `${run.fontSizePt}pt`,
    fontWeight: run.bold ? 'bold' : 'normal',
    fontStyle: run.italic ? 'italic' : 'normal',
    letterSpacing: run.letterSpacingPt === undefined ? undefined : `${run.letterSpacingPt}pt`,
    verticalAlign: run.baselineOffsetEm === undefined ? undefined : `${run.baselineOffsetEm}em`,
    textDecorationLine: [run.underline ? 'underline' : '', run.lineThrough ? 'line-through' : ''].filter(Boolean).join(' ') || 'none',
    textDecorationStyle: run.underlineStyle === 'wavy' ? 'wavy' : run.underlineStyle || run.lineThroughStyle,
    WebkitTextStroke: run.outlineColor !== undefined && run.outlineWidth !== undefined ? `${run.outlineWidth / 12700}pt ${run.outlineColor}` : undefined,
    textShadow: run.shadow ? `${Math.cos(angle) * distance}pt ${Math.sin(angle) * distance}pt ${run.shadow.blur / 12700}pt ${run.shadow.color}` : undefined,
  }
}
</script>

<template>
  <div class="speaker-notes-rich" data-presenter-notes>
    <p v-for="(entry, index) in paragraphEntries(paragraphs)" :key="index" :style="paragraphStyle(entry.paragraph)"><img v-if="entry.paragraph.bullet?.kind === 'image' && entry.paragraph.bullet.url" class="note-picture-bullet" :src="entry.paragraph.bullet.url" :style="bulletStyle(entry.paragraph)" alt=""><span v-else-if="entry.marker" class="note-bullet" :style="bulletStyle(entry.paragraph)">{{ entry.marker }}</span><component :is="run.hyperlink ? 'button' : 'span'" v-for="(run, runIndex) in entry.paragraph.runs" :key="runIndex" :lang="run.language" :style="runStyle(run)" :type="run.hyperlink ? 'button' : undefined" @click="run.hyperlink && emit('activate-hyperlink', run.hyperlink)"><span v-for="(segment, segmentIndex) in scriptSegments(run.text)" :key="segmentIndex" :style="{ fontFamily: segmentFont(run, segment.script) }">{{ segment.text }}</span></component>{{ index < paragraphs.length - 1 ? '\n' : '' }}</p>
  </div>
</template>

<style scoped>
.speaker-notes-rich { padding: 8px 15px 16px; overflow: auto; background: #fff; color: #000; font-size: 14px; line-height: 1.7; white-space: pre-wrap; overflow-wrap: anywhere; }
.speaker-notes-rich p { margin: 0; min-height: 1em; }
.note-bullet, .note-picture-bullet { display: inline-block; margin-right: .45em; }
.note-picture-bullet { width: 1em; height: 1em; object-fit: contain; vertical-align: middle; }
.speaker-notes-rich button { border: 0; padding: 0; background: transparent; font: inherit; color: inherit; white-space: inherit; cursor: pointer; }
.speaker-notes-rich button:focus-visible { outline: 2px solid #0563c1; outline-offset: 2px; }
</style>
