import type { PptxEmbeddedFontFace, PptxSlide, PptxTransitionPlayback, PptxTriggerPlayback } from '~/utils/pptx-model'

export const presenterChannelName = (id: string) => `deckline-presenter:${id}`

export interface PresenterPoint { x: number; y: number }
export interface PresenterInkStroke { tool: 'pen' | 'highlighter'; points: string }

export type PresenterMessage =
  | { type: 'ready' }
  | { type: 'document'; name: string; width: number; height: number; totalSlides: number; showAnimations: boolean; showNarration: boolean; embeddedFonts: PptxEmbeddedFontFace[]; fontWarnings: string[] }
  | { type: 'slide'; slide: PptxSlide; number: number; animationStep: number; blank: 'black' | 'white' | null; triggerPlayback: PptxTriggerPlayback[]; transitionPlayback: PptxTransitionPlayback | null; ink: PresenterInkStroke[]; draftStroke: PresenterInkStroke | null; laserPoint: PresenterPoint | null }
  | { type: 'state'; animationStep: number; blank: 'black' | 'white' | null; triggerPlayback: PptxTriggerPlayback[] }
  | { type: 'ink-start'; slideId: string; stroke: PresenterInkStroke }
  | { type: 'ink-point'; slideId: string; point: PresenterPoint }
  | { type: 'ink-end'; slideId: string }
  | { type: 'ink-clear'; slideId: string }
  | { type: 'laser-point'; slideId: string; point: PresenterPoint | null }
  | { type: 'clear' }
  | { type: 'rendered'; slideId: string; number: number; animationStep: number; blank: 'black' | 'white' | null; triggerCount: number }
  | { type: 'input'; action: 'next' | 'previous' | 'end' | 'blank-black' | 'blank-white' | 'blank-clear' | 'close' }
  | { type: 'input'; action: 'go-to'; slideNumber: number }
  | { type: 'closed' }
