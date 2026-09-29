import { DOMParser as XmlDomParser } from '@xmldom/xmldom'
import { parsePptx, type PptxDocument } from '../utils/pptx'

interface ParseRequest {
  id: number
  file: File
}

interface ParseReply {
  id: number
  document?: PptxDocument
  error?: string
}

class WorkerXmlParser extends XmlDomParser {
  constructor() {
    super({ onError: (_level, message) => { throw new Error(message) } })
  }
}

Object.defineProperty(globalThis, 'DOMParser', { configurable: true, value: WorkerXmlParser })

const scope = globalThis as unknown as {
  addEventListener(type: 'message', listener: (event: MessageEvent<ParseRequest>) => void): void
  postMessage(message: ParseReply): void
}

scope.addEventListener('message', event => {
  void parsePptx(event.data.file).then(
    document => scope.postMessage({ id: event.data.id, document }),
    error => scope.postMessage({
      id: event.data.id,
      error: error instanceof Error ? error.message : 'The presentation could not be read.',
    }),
  )
})
