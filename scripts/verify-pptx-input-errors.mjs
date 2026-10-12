import assert from 'node:assert/strict'
import { normalizePptxError, validatePptxFile } from '../app/utils/pptx-input.ts'

await assert.rejects(
  validatePptxFile({ size: Number.MAX_SAFE_INTEGER }),
  { message: 'This presentation is larger than the 150 MB browser limit.' }
)

const compoundFile = new File([Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])], 'protected.pptx')
await assert.rejects(
  validatePptxFile(compoundFile),
  { message: 'This file is encrypted or uses the legacy .ppt format. Save it as an unencrypted .pptx first.' }
)

assert.equal(normalizePptxError(new Error("Can't find end of central directory: is this a zip file?")).message, 'The PPTX file appears damaged or incomplete.')
assert.equal(normalizePptxError(new Error('Password required')).message, 'Password-protected presentations are not supported.')
assert.equal(normalizePptxError('unknown parser failure').message, 'The presentation could not be read.')

console.log('PPTX size, encryption, and archive error checks passed.')
