import type { PptxDocument } from './pptx-model'
import { registerPptxEmbeddedFonts, releasePptxFontFaces } from './pptx-fonts'

interface ParseReply {
  id: number
  document?: PptxDocument
  error?: string
}

let parserWorker: Worker | undefined
let nextRequestId = 0
const pending = new Map<number, { resolve: (document: PptxDocument) => void; reject: (error: Error) => void }>()
const loadedFonts = new WeakMap<PptxDocument, FontFace[]>()

function failWorker(error: Error) {
  parserWorker?.terminate()
  parserWorker = undefined
  for (const request of pending.values()) request.reject(error)
  pending.clear()
}

function getWorker(): Worker {
  if (parserWorker) return parserWorker
  parserWorker = new Worker(new URL('../workers/pptx-parser.worker.ts', import.meta.url), { type: 'module' })
  parserWorker.onmessage = ({ data }: MessageEvent<ParseReply>) => {
    const request = pending.get(data.id)
    if (!request) return
    pending.delete(data.id)
    if (data.error) request.reject(new Error(data.error))
    else if (data.document) request.resolve(data.document)
    else request.reject(new Error('The presentation parser returned no document.'))
  }
  parserWorker.onerror = event => {
    event.preventDefault()
    failWorker(new Error('The presentation parsing worker failed.'))
  }
  parserWorker.onmessageerror = () => failWorker(new Error('The presentation parser returned unreadable data.'))
  return parserWorker
}

function addMediaCapabilityWarnings(document: PptxDocument) {
  const warnings = new Set(document.warnings)
  const checked = new Set<string>()
  for (const element of document.slides.flatMap(slide => slide.elements)) {
    if (!element.mediaUrl || !element.mediaMime || !element.mediaType) continue
    const key = `${element.mediaType}:${element.mediaMime}`
    if (checked.has(key)) continue
    checked.add(key)
    if (!window.document.createElement(element.mediaType).canPlayType(element.mediaMime)) {
      warnings.add(`This browser may not support embedded ${element.mediaType} format ${element.mediaFormat || 'unknown'}.`)
    }
  }
  document.warnings = [...warnings]
}

async function loadEmbeddedFonts(presentation: PptxDocument) {
  const loaded = await registerPptxEmbeddedFonts(presentation.embeddedFonts)
  loadedFonts.set(presentation, loaded.faces)
  presentation.warnings = [...new Set([...presentation.warnings, ...loaded.warnings])]
}

export async function parsePptxInWorker(file: File): Promise<PptxDocument> {
  let document: PptxDocument
  let worker: Worker | undefined
  try {
    if (typeof Worker !== 'undefined') worker = getWorker()
  } catch {
    worker = undefined
  }
  if (!worker) {
    const { parsePptx } = await import('./pptx')
    document = await parsePptx(file)
    document.warnings = [...new Set([...document.warnings, 'The PPTX parsing worker could not start; parsing ran on the main thread and may be slower.'])]
  } else {
    const id = ++nextRequestId
    document = await new Promise<PptxDocument>((resolve, reject) => {
      pending.set(id, { resolve, reject })
      try {
        worker.postMessage({ id, file })
      } catch (error) {
        pending.delete(id)
        reject(error instanceof Error ? error : new Error('The presentation could not be sent to the parser.'))
      }
    })
  }
  addMediaCapabilityWarnings(document)
  await loadEmbeddedFonts(document)
  return document
}

export function disposePptxDocument(document: PptxDocument) {
  releasePptxFontFaces(loadedFonts.get(document) || [])
  loadedFonts.delete(document)
  document.objectUrls.forEach(URL.revokeObjectURL)
}

export function stopPptxParserWorker() {
  if (!parserWorker) return
  failWorker(new Error('The presentation parser was stopped.'))
}
