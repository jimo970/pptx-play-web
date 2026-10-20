import assert from 'node:assert/strict'
import { animationGeometryVariables, evaluateAnimationFormula, sampleAnimationFormula } from '../app/utils/animation-formula.ts'
import { inverseTimingProgress, parseAnimationSpeed, parseTimingFilter, timingProgress } from '../app/utils/animation-timing.ts'

const close = (actual, expected, message) => assert.ok(Math.abs(actual - expected) < 1e-9, `${message}: ${actual} != ${expected}`)
assert.equal(parseAnimationSpeed(null), 1, 'An omitted cTn speed must use normal playback speed.')
assert.equal(parseAnimationSpeed('200000'), 2, 'An OOXML fixed-point cTn speed must use thousandths of a percent.')
assert.equal(parseAnimationSpeed(' 200% '), 2, 'A percent-suffixed cTn speed must use its percent value.')
assert.equal(parseAnimationSpeed('-50%'), -0.5, 'A negative percent-suffixed cTn speed must retain reverse playback.')
assert.equal(parseAnimationSpeed('0%'), 0, 'A zero cTn speed remains available to the caller for invalid-speed fallback.')
assert.equal(parseAnimationSpeed('1e5'), undefined, 'Invalid cTn speed syntax must be rejected.')
assert.equal(parseAnimationSpeed(''), undefined, 'An empty cTn speed must be rejected.')
const bounceTimeFilter = [{input: 0, output: 0}, {input: 0.14, output: 0.36}, {input: 0.43, output: 0.73}, {input: 0.71, output: 0.91}, {input: 1, output: 1}]
close(timingProgress({timeFilter: bounceTimeFilter}, 0.14), 0.36, 'WPS Bounce time filter must remap progress')
assert.ok(Math.abs(inverseTimingProgress({timeFilter: bounceTimeFilter}, 0.36) - 0.14) < 1e-7)
assert.deepEqual(parseTimingFilter('0,0; 0.14,0.36; 0.43,0.73; 0.71,0.91; 1.0,1.0'), bounceTimeFilter)
close(timingProgress({timeFilter: bounceTimeFilter}, 0.25), 0.36 + 0.37 * 0.11 / 0.29, 'Time filter must interpolate between control points')
close(timingProgress({acceleration: 0.2, timeFilter: bounceTimeFilter}, 0.1), (0.1 ** 2 / (0.4 * 0.9)) * 0.36 / 0.14, 'Time filter follows acceleration mapping')
const reversingTimeFilter = [{input: 0, output: 0}, {input: 0.3, output: 0.8}, {input: 0.6, output: 0.2}, {input: 1, output: 1}]
assert.deepEqual(parseTimingFilter('0,0; 0.3,0.8; 0.6,0.2; 1,1'), reversingTimeFilter, 'A valid custom time warp may move local animation time backward before continuing.')
const reversingCurve = {timeFilter: reversingTimeFilter}
close(timingProgress(reversingCurve, 0.45), 0.5, 'Time filter must interpolate a descending progress segment')
close(inverseTimingProgress(reversingCurve, 0.5), 0.1875, 'Inverse time lookup must select the earliest crossing on a reversing curve')
close(timingProgress(reversingCurve, inverseTimingProgress(reversingCurve, 0.5)), 0.5, 'Inverse time lookup must return a point on the reversing curve')
const acceleratedReversingCurve = {acceleration: 0.2, timeFilter: reversingTimeFilter}
const acceleratedReversingProgress = timingProgress(acceleratedReversingCurve, 0.35)
close(inverseTimingProgress(acceleratedReversingCurve, acceleratedReversingProgress), 0.35, 'Inverse time lookup must undo acceleration before a reversing time filter')
const flashBulbTimeFilter = [{input: 0, output: 0}, {input: 0.2, output: 0.5}, {input: 0.8, output: 0.5}, {input: 1, output: 0}]
assert.deepEqual(parseTimingFilter('0, 0; .2, .5; .8, .5; 1, 0'), flashBulbTimeFilter, 'The WPS FlashBulb time warp returns to zero at its endpoint')
close(timingProgress({timeFilter: flashBulbTimeFilter}, 0.25), 0.5, 'FlashBulb must preserve its middle plateau')
close(timingProgress({timeFilter: flashBulbTimeFilter}, 1), 0, 'FlashBulb must hold the exported endpoint rather than force completion to one')
close(timingProgress({timeFilter: parseTimingFilter('0,1;1,0')}, 0), 1, 'A reversed time filter must preserve its nonzero starting point')
close(timingProgress({timeFilter: parseTimingFilter('0,0;1,0.5')}, 1), 0.5, 'A partial time filter must preserve its declared ending value')
for (const invalid of ['', '0,0;0,1;1,1', '0,0;0.5,0.8;0.5,0.4;1,1', '0,0;1,2', '0,0;1,', '0,0;0.5,NaN;1,1', '0.1,0;1,1', '0,0;0.9,1']) assert.equal(parseTimingFilter(invalid), undefined)
const geometry = animationGeometryVariables({x: 100, y: 200, width: 40, height: 60}, 1000, 500)
close(geometry.ppt_x, 0.12, 'ppt_x uses the shape center')
close(geometry.ppt_y, 0.46, 'ppt_y uses the shape center')
close((1 + geometry.ppt_h / 2 - geometry.ppt_y) * 500 + 200, 500, 'A bottom entrance begins exactly below the slide')
const transformed = animationGeometryVariables({x: 100, y: 200, width: 40, height: 60, renderMatrix: [0, 2, -2, 0, 120, 0]}, 1000, 500)
close(transformed.ppt_x, 0.16, 'Affine center includes the shape bounding-box offset')
close(transformed.ppt_y, 0.48, 'Affine center preserves rotation and scale')
close(transformed.ppt_w, 0.08, 'Shape width includes parent scale')
close(transformed.ppt_h, 0.24, 'Shape height includes parent scale')
// Microsoft's tav example and WPS Bounce use input values 0.5 -> 1, not 0 -> 1.
const bounceFrames = [{offset: 0, value: 0.5, formula: '#ppt_y-sin(pi*$)/3'}, {offset: 1, value: 1}]
const bounce = sampleAnimationFormula(bounceFrames, 664, {ppt_y: 0.6}, -100, 100)
assert.ok(bounce)
close(bounce[0].value, 0.6 - 1 / 3, 'Bounce must start above its original position')
close(bounce.at(-1).value, 0.6, 'Bounce must return to its original position')
close(bounce[30].value, 0.6 - Math.sin(0.75 * Math.PI) / 3, 'Formula input must interpolate keyframe values')
const discrete = sampleAnimationFormula(bounceFrames, 664, {ppt_y: 0.6}, -100, 100, 'discrete')
close(discrete[30].value, 0.6 - 1 / 3, 'Discrete formula input must retain its prior keyframe value')

// A narrow WPS Swivel still uses dimensionless input 0 -> 1 and a static width.
const swivelFrames = [{offset: 0, value: 0, formula: '#ppt_w*sin(2.5*pi*$)'}, {offset: 1, value: 1}]
const swivel = sampleAnimationFormula(swivelFrames, 2000, {ppt_w: 0.02}, -0.2, 0.2)
assert.ok(swivel)
close(swivel[60].value, -Math.SQRT1_2 * 0.02, 'Swivel must preserve its signed width')
close(swivel.at(-1).value, 0.02, 'Swivel must finish at the original width')

close(evaluateAnimationFormula('4*ppt_w', {ppt_w: 0.02}, true), 0.08, 'WPS strVal may omit the hash prefix')
assert.equal(evaluateAnimationFormula('4*ppt_w', {ppt_w: 0.02}), undefined, 'Standard formula grammar keeps its hash prefix')
assert.equal(evaluateAnimationFormula('constructor(1)', {}, true), undefined)
assert.equal(evaluateAnimationFormula('globalThis.process.exit()', {}, true), undefined)
assert.equal(evaluateAnimationFormula('#missing+1', {}), undefined)
assert.equal(evaluateAnimationFormula('1/0', {}), undefined)
assert.equal(evaluateAnimationFormula('1+'.repeat(513) + '1', {}), undefined)
assert.equal(sampleAnimationFormula([{offset: 0, value: 0, formula: '11*#ppt_w'}, {offset: 1, value: 1}], 200, {ppt_w: 0.02}, -0.2, 0.2), undefined)
assert.equal(sampleAnimationFormula([], 200, {}, 0, 1), undefined)
assert.equal(sampleAnimationFormula([{offset: 0.5, value: 0}, {offset: 0.5, value: 1}], 200, {}, 0, 1), undefined)
console.log('Animation formula input, static geometry, signed width, WPS value aliases, cTn speed forms, timing filters, inverse clocks, and invalid-value checks passed.')
