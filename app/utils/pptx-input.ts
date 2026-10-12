export const MAX_PPTX_FILE_BYTES = 150 * 1024 * 1024

const COMPOUND_FILE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const

export async function validatePptxFile(file: Pick<File, 'size' | 'slice'>): Promise<void> {
  if (file.size > MAX_PPTX_FILE_BYTES) throw new Error('This presentation is larger than the 150 MB browser limit.')
  const signature = new Uint8Array(await file.slice(0, 8).arrayBuffer())
  if (COMPOUND_FILE_SIGNATURE.every((byte, index) => signature[index] === byte)) {
    throw new Error('This file is encrypted or uses the legacy .ppt format. Save it as an unencrypted .pptx first.')
  }
}

export function normalizePptxError(error: unknown): Error {
  const message = error instanceof Error ? error.message : 'The presentation could not be read.'
  if (/encrypt|password/i.test(message)) return new Error('Password-protected presentations are not supported.')
  if (/zip|central directory|unexpected end/i.test(message)) return new Error('The PPTX file appears damaged or incomplete.')
  return error instanceof Error ? error : new Error(message)
}
