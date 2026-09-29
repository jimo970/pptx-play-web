import type { PptxSlide, PptxTransitionPlayback, PptxTriggerPlayback } from '~/utils/pptx'

export const presenterChannelName = (id: string) => `deckline-presenter:${id}`

export type PresenterMessage =
  | { type: 'ready' }
  | { type: 'document'; name: string; width: number; height: number; totalSlides: number }
  | { type: 'slide'; slide: PptxSlide; number: number; animationStep: number; blank: 'black' | 'white' | null; triggerPlayback: PptxTriggerPlayback[]; transitionPlayback: PptxTransitionPlayback | null }
  | { type: 'state'; animationStep: number; blank: 'black' | 'white' | null; triggerPlayback: PptxTriggerPlayback[] }
  | { type: 'clear' }
  | { type: 'rendered'; slideId: string; number: number; animationStep: number; blank: 'black' | 'white' | null; triggerCount: number }
  | { type: 'input'; action: 'next' | 'previous' | 'blank-black' | 'blank-white' | 'blank-clear' | 'close' }
  | { type: 'closed' }
