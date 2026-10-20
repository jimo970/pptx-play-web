import assert from 'node:assert/strict'
import {createTextIterationSegmenter, isTextIterationUnit, textIterationLocale, textIterationRanges} from '../app/utils/text-iteration.ts'

const locale = textIterationLocale([
  {text: '  ', language: 'en-US'},
  {text: 'ภาษาไทยกำลังเล่นทีละคำ', language: 'th-TH-u-ca-buddhist'}
])
assert.equal(locale, 'th-TH', 'Word iteration must use the source run language, stripped of unrelated locale extensions.')

const wordSegmenter = createTextIterationSegmenter(locale, 'word')
assert.equal(wordSegmenter.resolvedOptions().granularity, 'word')
assert.equal(wordSegmenter.resolvedOptions().locale.toLowerCase(), 'th-th', 'The requested word-segmentation locale must reach Intl.Segmenter.')

const graphemeLocale = textIterationLocale([{text: 'PPT', language: 'en-US'}])
assert.equal(createTextIterationSegmenter(undefined, 'grapheme').resolvedOptions().granularity, 'grapheme')
assert.equal(textIterationLocale([{text: 'PPT', language: 'en_US'}, {text: 'PPT', language: 'en-US'}]), 'en-US', 'Invalid tags must not prevent a later valid run language from being used.')
assert.equal(textIterationLocale([{text: 'PPT'}]), undefined, 'Untagged text must retain the runtime locale fallback.')
assert.equal(graphemeLocale, 'en-US')

assert.equal(isTextIterationUnit('lt', {segment: 'A'}), true, 'A letter must consume a per-letter iteration slot.')
assert.equal(isTextIterationUnit('lt', {segment: ' '}), false, 'A regular space must not consume a per-letter iteration slot.')
assert.equal(isTextIterationUnit('lt', {segment: '\u00a0'}), false, 'A non-breaking space must not consume a per-letter iteration slot.')
assert.equal(isTextIterationUnit('lt', {segment: 'e\u0301'}), true, 'A multi-code-point grapheme must remain one per-letter iteration slot.')
assert.equal(isTextIterationUnit('wd', {segment: 'word', isWordLike: true}), true, 'A word-like segment must consume a per-word iteration slot.')
assert.equal(isTextIterationUnit('wd', {segment: ' ', isWordLike: false}), false, 'A separator must not consume a per-word iteration slot.')
const letterRanges = [...createTextIterationSegmenter(undefined, 'grapheme').segment('A   B e\u0301')]
    .map(segment => ({segment: segment.segment, index: segment.index}))
assert.deepEqual(textIterationRanges('lt', letterRanges), [
    {start: 0, end: 1},
    {start: 4, end: 5},
    {start: 6, end: 8}
], 'Ignored whitespace must not shift the original UTF-16 character ranges.')

const wordSegments = createTextIterationSegmenter('en-US', 'word').segment('“Hello,”  world!')
assert.deepEqual(textIterationRanges('wd', wordSegments), [
    {start: 0, end: 8},
    {start: 10, end: 16}
], 'Punctuation adjacent to a word must share that word’s animation slot, while whitespace keeps its source offset.')
assert.deepEqual(textIterationRanges('wd', createTextIterationSegmenter(undefined, 'word').segment('one   two')),
    [{start: 0, end: 3}, {start: 6, end: 9}], 'Whitespace must not consume per-word iteration slots.')

console.log('Text iteration locale, whitespace, and word-punctuation slot selection passed.')
