import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'

const baseUrl = process.env.DECKLINE_BASE_URL || 'http://127.0.0.1:3000/'
const keepFixture = process.env.DECKLINE_KEEP_FIXTURE === '1'
const inputPptx = process.env.DECKLINE_INPUT_PPTX
const captureOnly = process.env.DECKLINE_CAPTURE_ONLY === '1'
const browserPath = process.env.DECKLINE_BROWSER || [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
].find(existsSync)

if (!browserPath) throw new Error('Set DECKLINE_BROWSER to an Edge or Chrome executable.')
try {
  const response = await fetch(baseUrl)
  if (!response.ok) throw new Error()
  await response.arrayBuffer()
} catch {
  throw new Error(`Start the player before this check: npm run dev -- --port 3000`)
}

const runtimeDir = await mkdtemp(join(tmpdir(), 'deckline-animation-'))
const fixturePath = join(runtimeDir, 'animation-fixture.pptx')
const profilePath = join(runtimeDir, 'browser-profile')
const debugPort = 12_000 + Math.floor(Math.random() * 2_000)
let browser
let socket
let closeBrowser = async () => {}

const delay = milliseconds => new Promise(resolve => setTimeout(resolve, milliseconds))
const assert = (condition, message) => { if (!condition) throw new Error(message) }

async function createFixture() {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide5.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide6.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide7.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide8.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide9.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide10.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide11.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide12.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide13.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide14.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide15.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>`)
  zip.file('ppt/presentation.xml', `<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId id="258" r:id="rId3"/><p:sldId id="259" r:id="rId4"/><p:sldId id="260" r:id="rId5"/><p:sldId id="261" r:id="rId6"/><p:sldId id="262" r:id="rId7"/><p:sldId id="263" r:id="rId8"/><p:sldId id="264" r:id="rId9"/><p:sldId id="265" r:id="rId10"/><p:sldId id="266" r:id="rId11"/><p:sldId id="267" r:id="rId12"/><p:sldId id="268" r:id="rId13"/><p:sldId id="269" r:id="rId14"/><p:sldId id="270" r:id="rId15"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`)
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide4.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide5.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide6.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide7.xml"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide8.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide9.xml"/><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide10.xml"/><Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide11.xml"/><Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide12.xml"/><Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide13.xml"/><Relationship Id="rId14" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide14.xml"/><Relationship Id="rId15" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide15.xml"/></Relationships>`)
  const shadow = `<a:effectLst><a:outerShdw blurRad="127000" dist="127000" dir="2700000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>`
  const shape = (id, name, x, y, color, width = 2600000, height = 2500000) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></p:spPr></p:sp>`
  const shadowedWithPrevious = shape(3, 'With previous', 3000000, 0, '2563eb').replace('</p:spPr>', `${shadow}</p:spPr>`)
  const customGeometryShape = `<p:sp><p:nvSpPr><p:cNvPr id="61" name="Custom geometry"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="9000000" y="6100000"/><a:ext cx="2600000" cy="600000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="0" b="0"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="500"/></a:moveTo><a:cubicBezTo><a:pt x="250" y="0"/><a:pt x="750" y="0"/><a:pt x="1000" y="500"/></a:cubicBezTo><a:cubicBezTo><a:pt x="750" y="1000"/><a:pt x="250" y="1000"/><a:pt x="0" y="500"/></a:cubicBezTo><a:close/></a:path><a:path w="1000" h="1000" fill="none"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:quadBezTo><a:pt x="500" y="1000"/><a:pt x="1000" y="0"/></a:quadBezTo><a:lnTo><a:pt x="1000" y="1000"/></a:lnTo><a:close/></a:path><a:path w="1000" h="1000" fill="none"><a:moveTo><a:pt x="100" y="500"/></a:moveTo><a:arcTo wR="400" hR="250" stAng="10800000" swAng="10800000"/><a:lnTo><a:pt x="100" y="500"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="a3e635"/></a:solidFill><a:ln w="25400" cap="rnd"><a:solidFill><a:srgbClr val="166534"/></a:solidFill><a:custDash><a:ds d="250000" sp="50000"/><a:ds d="50000" sp="25000"/></a:custDash></a:ln></p:spPr></p:sp>`
  const placeholderShape = `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Inherited paragraph"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="2"/><a:r><a:t>Inherited placeholder</a:t></a:r></a:p></p:txBody></p:sp>`
  const mediaPicture = (id, name, x, y, tag, relation, mediaOptions = '', mediaLink) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr><a:${tag} r:link="${relation}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media${mediaLink ? ` r:link="${mediaLink}"` : ''} r:embed="${relation}">${mediaOptions}</p14:media></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="5000000" cy="2800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  const picture = (id, name, x) => `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdBullet"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="4900000"/><a:ext cx="800000" cy="700000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  const lineShape = (id, name, x, y, lineEnds = '') => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="2200000" cy="500000"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="63500"><a:solidFill><a:srgbClr val="0000ff"/></a:solidFill>${lineEnds}</a:ln></p:spPr></p:sp>`
  const gradientShape = `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Gradient shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="4000000"/><a:ext cx="2600000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ef4444"/></a:gs><a:gs pos="100000"><a:srgbClr val="2563eb"><a:alpha val="50000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill>${shadow}</p:spPr></p:sp>`
  const luminanceShape = `<p:sp><p:nvSpPr><p:cNvPr id="5" name="Luminance transform"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="3000000" y="4000000"/><a:ext cx="2600000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="00ff00"><a:lumMod val="50000"/><a:lumOff val="20000"/></a:srgbClr></a:solidFill></p:spPr></p:sp>`
  const nestedGroup = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Nested group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm flipH="1"><a:off x="600000" y="300000"/><a:ext cx="600000" cy="400000"/><a:chOff x="100000" y="50000"/><a:chExt cx="300000" cy="200000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="11" name="Nested group child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="50000"/><a:ext cx="300000" cy="200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0ea5e9"/></a:solidFill></p:spPr></p:sp></p:grpSp>`
  const groupedShape = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="Rotated group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000"><a:off x="5500000" y="3800000"/><a:ext cx="4000000" cy="2000000"/><a:chOff x="100000" y="200000"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="9" name="Rotated group child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="200000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="f43f5e"/></a:solidFill></p:spPr></p:sp>${nestedGroup}</p:grpSp>`
  const parentMotionGroup = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="124" name="Parent motion group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="2000000"/><a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="125" name="Parent origin grouped motion path"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0d9488"/></a:solidFill></p:spPr></p:sp></p:grpSp>`
  const elementIterationGroup = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="62" name="Element iteration group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="8500000" y="4500000"/><a:ext cx="3200000" cy="1500000"/><a:chOff x="0" y="0"/><a:chExt cx="3200000" cy="1500000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="63" name="Element iteration first"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1400000" cy="1400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0ea5e9"/></a:solidFill></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="64" name="Element iteration second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1800000" y="0"/><a:ext cx="1400000" cy="1400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="8b5cf6"/></a:solidFill></p:spPr></p:sp></p:grpSp>`
  const textShape = (id, name, x, y) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="2600000" cy="2500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="e2e8f0"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800">${shadow}</a:rPr><a:t>Built paragraph</a:t></a:r></a:p><a:p><a:r><a:rPr sz="1800"/><a:t>Al</a:t></a:r><a:br/><a:r><a:rPr sz="1800" b="1"/><a:t>ways visible</a:t></a:r></a:p></p:txBody></p:sp>`
const guideDrivenGeometryPath = `<a:path w="1000" h="1000"><a:moveTo><a:pt x="xQuarter" y="bottom"/></a:moveTo><a:lnTo><a:pt x="xThreeQuarter" y="bottom"/></a:lnTo><a:lnTo><a:pt x="xThreeQuarter" y="top"/></a:lnTo><a:close/></a:path>`
const guideFormulaPath = `<a:path w="1000" h="1000"><a:moveTo><a:pt x="addDiv" y="ifElse"/></a:moveTo><a:lnTo><a:pt x="absVal" y="maxVal"/><a:pt x="minVal" y="modVal"/><a:pt x="catVal" y="cosVal"/><a:pt x="ifElseFallback" y="satVal"/><a:pt x="sinVal" y="sqrtVal"/><a:pt x="tanVal" y="pinLow"/><a:pt x="pinHigh" y="adjX"/></a:lnTo></a:path>`
const guideFormulaArcPath = `<a:path w="1000" h="1000"><a:moveTo><a:pt x="300" y="500"/></a:moveTo><a:arcTo wR="200" hR="200" stAng="angle45" swAng="cd4"/></a:path>`
const customGeometryShapeWithGuides = customGeometryShape
  .replace('<a:avLst/><a:gdLst/>', '<a:avLst><a:gd name="adjX" fmla="val 250"/></a:avLst><a:gdLst><a:gd name="xQuarter" fmla="*/ w adjX 1000"/><a:gd name="xThreeQuarter" fmla="+- w 0 xQuarter"/><a:gd name="bottom" fmla="pin 0 750 h"/><a:gd name="top" fmla="min h 250"/><a:gd name="addDiv" fmla="+/ 300 100 2"/><a:gd name="ifElse" fmla="?: 1 50 0"/><a:gd name="ifElseFallback" fmla="?: 0 50 25"/><a:gd name="absVal" fmla="abs -10"/><a:gd name="angle45" fmla="at2 1 1"/><a:gd name="catVal" fmla="cat2 100 1 0"/><a:gd name="cosVal" fmla="cos 100 0"/><a:gd name="maxVal" fmla="max 10 20"/><a:gd name="minVal" fmla="min 10 20"/><a:gd name="modVal" fmla="mod 3 4 12"/><a:gd name="satVal" fmla="sat2 100 1 0"/><a:gd name="sinVal" fmla="sin 100 0"/><a:gd name="sqrtVal" fmla="sqrt 256"/><a:gd name="tanVal" fmla="tan 100 0"/><a:gd name="pinLow" fmla="pin 0 -5 1000"/><a:gd name="pinHigh" fmla="pin 0 1200 1000"/></a:gdLst>')
  .replace('</a:pathLst>', `${guideDrivenGeometryPath}${guideFormulaPath}${guideFormulaArcPath}</a:pathLst>`)
const customGeometryTextShape = `<p:sp><p:nvSpPr><p:cNvPr id="65" name="Custom text rectangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="5000000" y="4200000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst><a:gd name="quarterWidth" fmla="*/ w 1 4"/></a:gdLst><a:ahLst/><a:cxnLst/><a:rect l="quarterWidth" t="200" r="900" b="800"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="500"/></a:moveTo><a:lnTo><a:pt x="1000" y="500"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="38bdf8"/></a:solidFill></p:spPr><p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Inside custom text rectangle</a:t></a:r></a:p></p:txBody></p:sp>`
const invalidCustomGeometryTextShape = `<p:sp><p:nvSpPr><p:cNvPr id="66" name="Invalid custom text rectangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="7500000" y="4200000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="800" t="200" r="200" b="800"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="500"/></a:moveTo><a:lnTo><a:pt x="1000" y="500"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="f97316"/></a:solidFill></p:spPr><p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Fallback text box</a:t></a:r></a:p></p:txBody></p:sp>`
const customGradientShape = `<p:sp><p:nvSpPr><p:cNvPr id="67" name="Custom gradient geometry"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10000000" y="4200000"/><a:ext cx="1500000" cy="1000000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="1000" b="1000"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="1000" y="0"/><a:pt x="1000" y="1000"/><a:pt x="0" y="1000"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ff0000"/></a:gs><a:gs pos="50000"><a:srgbClr val="00ff00"><a:alpha val="50000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="0000ff"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill><a:ln w="63500"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ffff00"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></a:ln></p:spPr></p:sp>`
const customRadialGradientShape = `<p:sp><p:nvSpPr><p:cNvPr id="68" name="Custom radial gradient geometry"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="9500000" y="5200000"/><a:ext cx="1500000" cy="1000000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="1000" b="1000"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="1000" y="0"/><a:pt x="1000" y="1000"/><a:pt x="0" y="1000"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ff0000"><a:alpha val="50000"/></a:srgbClr></a:gs><a:gs pos="50000"><a:srgbClr val="00ff00"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000ff"/></a:gs></a:gsLst><a:path path="circle"/></a:gradFill><a:ln w="63500"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ffff00"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst><a:path path="circle"/></a:gradFill></a:ln></p:spPr></p:sp>`
const characterSpacedShapeId = 59
const baselineShiftedShapeId = 45
const characterSlide = `<p:par><p:cTn id="18001" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="slide(fromLeft)"><p:cBhvr><p:cTn id="18002" dur="400" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="160"><p:txEl><p:charRg st="4" end="8"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
const plainTextShape = (id, name, x, y, text, outline = '') => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="2600000" cy="800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="e2e8f0"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"${id === characterSpacedShapeId ? ' spc="250"' : ''}${id === baselineShiftedShapeId ? ' baseline="30000"' : ''}>${outline}</a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
  const autoFitText = (id, name, x, y, width, height, size, text) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" bIns="0" anchor="t"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="${size}" b="1"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
  const tableCell = (text, attributes = '', color = 'ffffff') => `<a:tc ${attributes}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800" b="1"/><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr marL="91440" marR="91440" marT="45720" marB="45720"><a:lnL w="12700"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:lnR><a:lnT w="12700"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:lnT><a:lnB w="12700"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:lnB><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:tcPr></a:tc>`
  const autofitTableCell = tableCell('Table cell normal autofit shrinks overflowing text to stay inside this saved row. '.repeat(4), '', 'dbeafe').replace('<a:bodyPr/>', '<a:bodyPr><a:normAutofit/></a:bodyPr>')
  const columnTableCell = tableCell('This cell fills its first text column, then overflows into the left column with a gap.', '', 'fef3c7').replace('<a:bodyPr/>', '<a:bodyPr numCol="2" spcCol="457200" rtlCol="1"/>')
  const table = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Basic table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="3000000" y="500000"/><a:ext cx="6000000" cy="4500000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="3000000"/><a:gridCol w="3000000"/></a:tblGrid><a:tr h="1500000">${tableCell('Merged header', 'gridSpan="2"', '1e40af')}${tableCell('', 'hMerge="1"')}</a:tr><a:tr h="1500000">${tableCell('Left cell')}${autofitTableCell}</a:tr><a:tr h="1500000">${columnTableCell}${tableCell('Next cell')}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>` + customRadialGradientShape
  const chartFrame = (id, name, relationship, x, y, width, height) => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${relationship}"/></a:graphicData></a:graphic></p:graphicFrame>`
  const chartPart = `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarterly sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>North</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="0ea5e9"/></a:solidFill></c:spPr><c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>18</c:v></c:pt><c:pt idx="2"><c:v>26</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>South</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="f97316"/></a:solidFill></c:spPr><c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>9</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt><c:pt idx="2"><c:v>21</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart></c:plotArea></c:chart></c:chartSpace>`
  const tree = body => `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${body}</p:spTree></p:cSld>`
  const silentWave = (durationSeconds = 0.5) => {
    const sampleRate = 8000, samples = sampleRate * durationSeconds
    const header = Buffer.alloc(44)
    header.write('RIFF', 0); header.writeUInt32LE(36 + samples * 2, 4); header.write('WAVEfmt ', 8); header.writeUInt32LE(16, 16)
    header.writeUInt16LE(1, 20); header.writeUInt16LE(1, 22); header.writeUInt32LE(sampleRate, 24); header.writeUInt32LE(sampleRate * 2, 28)
    header.writeUInt16LE(2, 32); header.writeUInt16LE(16, 34); header.write('data', 36); header.writeUInt32LE(samples * 2, 40)
    return Buffer.concat([header, Buffer.alloc(samples * 2)])
  }
  const effectGroup = (id, nodeType, target, transition, filter, animationId, duration, groupDelay = '0', paragraph, timingAttributes = '', fill = 'hold', groupTimingAttributes = '', childDelay = '0') => `<p:par><p:cTn id="${id}" fill="hold"${groupTimingAttributes} nodeType="${nodeType}"><p:stCondLst><p:cond delay="${nodeType === 'clickEffect' ? 'indefinite' : groupDelay}"/></p:stCondLst><p:childTnLst><p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr><p:cTn id="${animationId}" dur="${duration}" fill="${fill}"${timingAttributes}><p:stCondLst><p:cond delay="${childDelay}"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${target}">${paragraph === undefined ? '' : `<p:txEl><p:pRg st="${paragraph}" end="${paragraph}"/></p:txEl>`}</p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const emptyMotionGroup = `<p:par><p:cTn id="150" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animMotion><p:cBhvr><p:cTn id="151" dur="1000" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="6"/></p:tgtEl></p:cBhvr></p:animMotion></p:childTnLst></p:cTn></p:par>`
  const characterFade = `<p:par><p:cTn id="91" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="92" dur="250" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="7"><p:txEl><p:charRg st="16" end="23"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect><p:animEffect transition="in" filter="wipe(right)"><p:cBhvr><p:cTn id="134" dur="250" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="59"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const characterColor = `<p:par><p:cTn id="93" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animClr><p:cBhvr><p:cTn id="94" dur="250" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="7"><p:txEl><p:charRg st="16" end="23"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="000000"/></p:from><p:to><a:srgbClr val="ff0000"/></p:to></p:animClr></p:childTnLst></p:cTn></p:par>`
  const characterFontSize = `<p:par><p:cTn id="95" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:anim to="1.5" calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="96" dur="250" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="7"><p:txEl><p:charRg st="19" end="23"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.fontSize</p:attrName></p:attrNameLst></p:cBhvr></p:anim></p:childTnLst></p:cTn></p:par>`
  const letterFade = `<p:par><p:cTn id="97" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:iterate type="lt"><p:tmAbs val="150"/></p:iterate><p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="98" dur="200" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="43"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const wordFade = `<p:par><p:cTn id="101" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:iterate type="wd" backwards="1"><p:tmPct val="50000"/></p:iterate><p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="102" dur="200" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="45"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const objectBlinds = `<p:par><p:cTn id="105" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="blinds(horizontal)"><p:cBhvr><p:cTn id="106" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="46"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const objectChecker = effectGroup(107, 'withEffect', 47, 'in', 'checkerboard(across)', 108, 400)
  const parentEasedGroup = effectGroup(258, 'withEffect', 102, 'in', 'fade', 259, 300, '0', undefined, '', 'hold', ' dur="600" accel="50000" decel="50000"')
  const parentSpeedGroup = effectGroup(262, 'withEffect', 104, 'in', 'fade', 263, 300, '0', undefined, '', 'hold', ' dur="600" spd="200000"', '75')
  const parentTimingAfterEffect = effectGroup(260, 'afterEffect', 103, 'in', 'fade', 261, 100)
  const parentRepeatCount = effectGroup(300, 'clickEffect', 106, 'in', 'fade', 301, 300, '0', undefined, '', 'hold', ' dur="600" repeatCount="2000" repeatDur="indefinite"')
  const parentRepeatDuration = effectGroup(303, 'clickEffect', 108, 'in', 'fade', 304, 300, '0', undefined, '', 'hold', ' dur="600" repeatCount="indefinite" repeatDur="900"')
  const parentRepeatAfterEffect = effectGroup(305, 'afterEffect', 109, 'in', 'fade', 306, 100)
  const parentReverseSpeedGroup = effectGroup(800, 'clickEffect', 113, 'in', 'fade', 801, 600, '0', undefined, '', 'hold', ' dur="600" spd="-200000"')
  const parentAutoReverseGroup = effectGroup(802, 'clickEffect', 115, 'in', 'fade', 803, 600, '0', undefined, '', 'hold', ' dur="600" autoRev="1"')
  const parentIndefiniteRepeat = effectGroup(807, 'clickEffect', 117, 'in', 'fade', 808, 180, '0', undefined, '', 'hold', ' dur="360" repeatDur="indefinite"')
  const childRepeatDuration = effectGroup(810, 'clickEffect', 119, 'in', 'fade', 811, 150, '0', undefined, ' repeatDur="450"')
  const triggeredSequence = (id, trigger, groups) => `<p:seq><p:cTn id="${id}" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${trigger}"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${groups}</p:childTnLst></p:cTn></p:seq>`
  const parentRepeatSequences = `<p:seq><p:cTn id="307" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="105"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${parentRepeatCount}</p:childTnLst></p:cTn></p:seq><p:seq><p:cTn id="308" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="107"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${parentRepeatDuration}${parentRepeatAfterEffect}</p:childTnLst></p:cTn></p:seq>${triggeredSequence(804, 112, parentReverseSpeedGroup)}${triggeredSequence(805, 114, parentAutoReverseGroup)}${triggeredSequence(806, 116, parentIndefiniteRepeat)}${triggeredSequence(809, 118, childRepeatDuration)}`
  const objectShapeElements = `${shape(49, 'Circle mask animation', 0, 6500000, '0ea5e9')}${shape(50, 'Diamond mask animation', 3000000, 6500000, '3b82f6')}${shape(51, 'Box mask animation', 6000000, 6500000, '2563eb')}${shape(52, 'Dissolve mask animation', 0, 3500000, '0891b2')}${shape(53, 'Wheel mask animation', 3000000, 3500000, '2563eb')}${shape(54, 'Wedge mask animation', 6000000, 3500000, '7c3aed')}${shape(55, 'Plus in mask animation', 0, 8500000, '0891b2')}${shape(56, 'Plus out mask animation', 6000000, 8500000, '7c3aed')}${shape(57, 'Random horizontal bars animation', 0, 10500000, '0284c7')}${shape(58, 'Random vertical bars animation', 6000000, 10500000, '7c3aed')}${shape(74, 'Media after-effect', 9000000, 7000000, '10b981')}`
  const shapeMaskClick = `${effectGroup(111, 'clickEffect', 49, 'in', 'circle(in)', 112, 400)}${effectGroup(113, 'clickEffect', 50, 'out', 'diamond(out)', 114, 400)}${effectGroup(115, 'clickEffect', 51, 'in', 'box(in)', 116, 400)}${effectGroup(117, 'clickEffect', 52, 'in', 'dissolve', 118, 400)}${effectGroup(119, 'clickEffect', 53, 'in', 'wheel(4)', 120, 400)}${effectGroup(121, 'clickEffect', 54, 'in', 'wedge', 122, 400)}${effectGroup(123, 'clickEffect', 55, 'in', 'plus(in)', 124, 400)}${effectGroup(125, 'clickEffect', 56, 'out', 'plus(out)', 126, 400)}${effectGroup(127, 'clickEffect', 57, 'in', 'randombar(horizontal)', 128, 400)}${effectGroup(129, 'clickEffect', 58, 'in', 'randombar(vertical)', 130, 400)}`
  const objectStripsElements = ['downLeft', 'upLeft', 'downRight', 'upRight'].map((direction, index) => shape(65 + index, `Strips ${direction} animation`, index * 3000000, 4200000, 'f97316')).join('')
  const objectStripsClick = ['downLeft', 'upLeft', 'downRight', 'upRight'].map((direction, index) => effectGroup(131 + index * 2, 'clickEffect', 65 + index, 'in', `strips(${direction})`, 132 + index * 2, 400)).join('')
  const barnDirections = ['inHorizontal', 'outHorizontal', 'inVertical', 'outVertical']
  const objectBarnElements = barnDirections.map((direction, index) => shape(69 + index, `Barn ${direction} animation`, index * 3000000, 5600000, 'a855f7')).join('')
  const objectBarnClick = barnDirections.map((direction, index) => {
    const motion = direction.startsWith('in') ? 'in' : 'out'
    return effectGroup(139 + index * 2, 'clickEffect', 69 + index, motion, `barn(${direction})`, 140 + index * 2, 400)
  }).join('')
  const objectCharacterWipe = `<p:par><p:cTn id="147" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="wipe(right)"><p:cBhvr><p:cTn id="148" dur="400" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="73"><p:txEl><p:charRg st="1" end="3"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const automaticMediaPlayback = `<p:par><p:cTn id="152" dur="indefinite" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:audio><p:cMediaNode vol="80000" mute="0"><p:cTn id="153" dur="indefinite"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cMediaNode></p:audio><p:video><p:cMediaNode numSld="2" vol="100000" mute="0"><p:cTn id="156" dur="100"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cMediaNode></p:video></p:childTnLst></p:cTn></p:par>`
  const mediaAfterEffect = effectGroup(163, 'afterEffect', 74, 'in', 'fade', 164, 150)
  const mediaCommandClick = (groupId, actionId, command) => `<p:par><p:cTn id="${groupId}" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="${command}"><p:cBhvr><p:cTn id="${actionId}" dur="0"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`
  const playFromClick = mediaCommandClick(154, 155, 'PlayFrom(0.2)')
  const togglePauseClick = mediaCommandClick(157, 158, 'togglePause')
  const pauseClick = mediaCommandClick(159, 160, 'pause')
  const stopClick = mediaCommandClick(161, 162, 'stop')
  const crossSlidePlaybackClick = `<p:par><p:cTn id="167" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:audio><p:cMediaNode numSld="2" vol="100000" mute="0"><p:cTn id="168" dur="indefinite"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="5"/></p:tgtEl></p:cMediaNode></p:audio></p:childTnLst></p:cTn></p:par>`
  const shapeMaskTiming = `<p:timing><p:tnLst><p:par><p:cTn id="120" dur="indefinite" nodeType="tmRoot"><p:childTnLst><p:seq><p:cTn id="121" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${automaticMediaPlayback}${mediaAfterEffect}${shapeMaskClick}${objectStripsClick}${objectBarnClick}${objectCharacterWipe}${playFromClick}${togglePauseClick}${pauseClick}${stopClick}${crossSlidePlaybackClick}</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
  const colorEffect = (id, target, property, from, to, paragraphRange = '', colorSpace = 'rgb', direction = 'cw') => `<p:par><p:cTn id="${id}" dur="indefinite" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animClr clrSpc="${colorSpace}" dir="${direction}"><p:cBhvr><p:cTn id="${id + 1}" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${target}">${paragraphRange ? `<p:txEl><p:pRg st="${paragraphRange.split(',')[0]}" end="${paragraphRange.split(',')[1]}"/></p:txEl>` : ''}</p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>${property}</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="${from}"/></p:from><p:to><a:srgbClr val="${to}"/></p:to></p:animClr></p:childTnLst></p:cTn></p:par>`
  const scaleEffect = `<p:par><p:cTn id="15" dur="indefinite" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animScale><p:cBhvr><p:cTn id="16" dur="200" fill="remove" repeatCount="1500" repeatDur="1000" autoRev="1"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="8"/></p:tgtEl></p:cBhvr><p:by x="150000" y="150000"/></p:animScale></p:childTnLst></p:cTn></p:par>`
  const motionEffect = `<p:par><p:cTn id="31" dur="indefinite" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animMotion origin="layout" path="M 0 0 L 0.001 0 M 0.2 0 L 0.25 0 E"><p:cBhvr><p:cTn id="32" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="30"/></p:tgtEl></p:cBhvr></p:animMotion></p:childTnLst></p:cTn></p:par>`
  const curveMotionEffect = `<p:par><p:cTn id="33" dur="indefinite" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animMotion origin="layout" path="M 0 0 C 0.1 0 0.15 0.2 0.25 0.2 c 0.1 0 0.15 -0.2 0.25 -0.2 Z E"><p:cBhvr><p:cTn id="34" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="29"/></p:tgtEl></p:cBhvr></p:animMotion></p:childTnLst></p:cTn></p:par>`
  const rotationEffect = `<p:par><p:cTn id="75" dur="indefinite" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animRot by="5400000"><p:cBhvr><p:cTn id="76" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="8"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>`
  const triggeredRotation = `<p:par><p:cTn id="73" dur="indefinite" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animRot by="3000000"><p:cBhvr><p:cTn id="74" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="9"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>`
  const shadowColorEffect = colorEffect(87, 3, 'shadow.color', '000000', 'ff0000').replace('<a:srgbClr val="000000"/>', '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>').replace('<a:srgbClr val="ff0000"/>', '<a:srgbClr val="ff0000"><a:alpha val="50000"/></a:srgbClr>')
  const triggeredColors = `${colorEffect(81, 3, 'fillcolor', '2563eb', 'ff0000')}${colorEffect(83, 42, 'stroke.color', '0000ff', 'ff0000', '', 'hsl', 'cw')}${colorEffect(85, 7, 'style.color', '000000', 'ff00ff', '0,0')}${shadowColorEffect}`
  const opacityEffect = (id, target, property, from, to) => `<p:par><p:cTn id="${id}" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:anim from="${from}" to="${to}" calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="${id + 1}" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${target}"/></p:tgtEl><p:attrNameLst><p:attrName>${property}</p:attrName></p:attrNameLst></p:cBhvr></p:anim></p:childTnLst></p:cTn></p:par>`
  const keyframedOpacity = `<p:par><p:cTn id="276" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="277" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="42"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0.2"/></p:val></p:tav><p:tav tm="50000"><p:val><p:fltVal val="0.8"/></p:val></p:tav><p:tav tm="100000"><p:val><p:fltVal val="0.4"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par>`
  const discreteOpacity = keyframedOpacity.replace('id="276"', 'id="278"').replace('id="277"', 'id="279"').replace('calcmode="lin"', 'calcmode="discrete"').replace('spid="42"', 'spid="8"').replace('val="0.2"', 'val="0.3"')
  const discreteFontSize = keyframedOpacity.replace('id="276"', 'id="280"').replace('id="277"', 'id="281"').replace('calcmode="lin"', 'calcmode="discrete"').replace('spid="42"', 'spid="59"').replace('<p:attrName>style.opacity</p:attrName>', '<p:attrName>style.fontSize</p:attrName>').replace('val="0.2"', 'val="0.5"').replace('val="0.8"', 'val="2"').replace('val="0.4"', 'val="0.75"')
  const fontWeightEffect = opacityEffect(12000, 59, 'style.fontWeight', 400, 900)
  const keyframedFontWeight = keyframedOpacity.replace('id="276"', 'id="12002"').replace('id="277"', 'id="12003"').replace('spid="42"', 'spid="43"').replace('<p:attrName>style.opacity</p:attrName>', '<p:attrName>style.fontWeight</p:attrName>').replace('val="0.2"', 'val="400"').replace('val="0.8"', 'val="700"').replace('val="0.4"', 'val="900"')
  const formulaFontWeight = `<p:par><p:cTn id="12004" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="12005" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="45"/></p:tgtEl><p:attrNameLst><p:attrName>style.fontWeight</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0" fmla="#style.fontWeight+200*sin(pi*$)"><p:val><p:fltVal val="400"/></p:val></p:tav><p:tav tm="100000"><p:val><p:fltVal val="700"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par>`
  const characterRangeOpacity = keyframedOpacity.replace('id="276"', 'id="693"').replace('id="277"', 'id="694"').replace('<p:spTgt spid="42"/>', '<p:spTgt spid="670"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt>').match(/<p:anim[\s\S]*?<\/p:anim>/)?.[0]
  const setCharacterRangeOpacity = '<p:set><p:cBhvr><p:cTn id="696" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="2" end="4"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:fltVal val="0.35"/></p:to></p:set>'
  const opacityEffects = `${opacityEffect(198, 99, 'style.opacity', 1, 0.25)}${opacityEffect(200, 3, 'fill.opacity', 1, 0.5)}${opacityEffect(202, 42, 'stroke.opacity', 1, 0.5)}${opacityEffect(204, 3, 'shadow.opacity', 1, 0.5)}`
  const setOpacity = `<p:par><p:cTn id="270" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="271" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="61"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:fltVal val="0.5"/></p:to></p:set></p:childTnLst></p:cTn></p:par>`
  const setOpacityString = `<p:par><p:cTn id="272" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="273" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="59"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="0.75"/></p:to></p:set></p:childTnLst></p:cTn></p:par>`
  const setColor = `<p:par><p:cTn id="274" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="275" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="61"/></p:tgtEl><p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:clrVal><a:srgbClr val="ff0000"/></p:clrVal></p:to></p:set></p:childTnLst></p:cTn></p:par>`
  const unsupportedSet = `<p:par><p:cTn id="264" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="265" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="99"/></p:tgtEl><p:attrNameLst><p:attrName>style.fontSize</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:fltVal val="0.5"/></p:to></p:set></p:childTnLst></p:cTn></p:par>`
  const fontSizeSet = `<p:par><p:cTn id="286" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="287" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="59"><p:txEl><p:charRg st="15" end="21"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.fontSize</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:fltVal val="1.5"/></p:to></p:set></p:childTnLst></p:cTn></p:par>`
  const triggeredEffect = `<p:seq><p:cTn id="70" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="8"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${effectGroup(71, 'clickEffect', 2, 'in', 'fade', 72, 300)}${triggeredRotation}${triggeredColors}${opacityEffects}${fontWeightEffect}${keyframedFontWeight}${formulaFontWeight}${keyframedOpacity}${discreteOpacity}${discreteFontSize}${setOpacity}${setOpacityString}${setColor}${effectGroup(161, 'withEffect', 60, 'in', 'slide(fromLeft)', 162, 400)}${effectGroup(163, 'clickEffect', 60, 'out', 'slide(fromRight)', 164, 300)}</p:childTnLst></p:cTn></p:seq>`
  const triggeredMouseOver = `<p:seq><p:cTn id="180" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onMouseOver" delay="0"><p:tgtEl><p:spTgt spid="8"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${effectGroup(181, 'clickEffect', 5, 'in', 'fade', 182, 200)}</p:childTnLst></p:cTn></p:seq>`
  const triggeredMouseOut = `<p:seq><p:cTn id="183" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onMouseOut" delay="0"><p:tgtEl><p:spTgt spid="8"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${effectGroup(184, 'clickEffect', 5, 'out', 'fade', 185, 200)}</p:childTnLst></p:cTn></p:seq>`
  const triggeredDoubleClick = `<p:seq><p:cTn id="186" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onDblClick" delay="0"><p:tgtEl><p:spTgt spid="11"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${effectGroup(187, 'clickEffect', 5, 'in', 'fade', 188, 200)}</p:childTnLst></p:cTn></p:seq>`
  const triggeredSlideNext = `<p:seq><p:cTn id="192" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${effectGroup(193, 'clickEffect', 90, 'in', 'fade', 194, 300)}</p:childTnLst></p:cTn></p:seq>`
  const triggeredSlidePrevious = `<p:seq><p:cTn id="195" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${effectGroup(196, 'clickEffect', 91, 'in', 'fade', 197, 300)}</p:childTnLst></p:cTn></p:seq>`
  const speedTriggerSequence = `<p:seq><p:cTn id="250" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="98"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${effectGroup(251, 'clickEffect', 99, 'in', 'fade', 252, 300, '0', undefined, ' spd="200000"')}</p:childTnLst></p:cTn></p:seq>`
  const reverseSpeedSequence = `<p:seq><p:cTn id="253" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="100"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="254" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:animScale><p:cBhvr><p:cTn id="255" dur="300" fill="hold" spd="-200000"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="101"/></p:tgtEl></p:cBhvr><p:from x="100000" y="100000"/><p:to x="200000" y="200000"/></p:animScale></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq>`
  const elementIterationEffect = `<p:seq><p:cTn id="140" dur="indefinite" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="63"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="141" dur="indefinite" nodeType="clickEffect"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:iterate type="el"><p:tmAbs val="200"/></p:iterate><p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="142" dur="250" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="62"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq>`
  const wipeAnimations = ['right', 'left', 'up', 'down'].map((direction, index) => effectGroup(25 + index * 2, 'withEffect', 20 + index, 'in', `wipe(${direction})`, 26 + index * 2, 200, '0')).join('')
  const timing = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek" prevAc="skipTimed"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${fontSizeSet}${unsupportedSet}${emptyMotionGroup}${effectGroup(3, 'afterEffect', 6, 'in', 'fade', 4, 150, '0', undefined, ' repeatCount="2000"', 'remove')}${effectGroup(5, 'clickEffect', 2, 'in', 'fade', 6, 300, '0', undefined, ' accel="50000" decel="50000"')}${parentSpeedGroup}${motionEffect}${curveMotionEffect}${effectGroup(17, 'withEffect', 24, 'out', 'wipe(left)', 18, 200)}${effectGroup(7, 'withEffect', 3, 'in', 'fade', 8, 200, '100')}${effectGroup(9, 'afterEffect', 4, 'out', 'fade', 10, 200, '50', undefined, '', 'freeze')}${characterFade}${characterColor}${characterFontSize}${letterFade}${wordFade}${objectBlinds}${objectChecker}${effectGroup(103, 'afterEffect', 44, 'out', 'fade', 104, 100)}${parentEasedGroup}${parentTimingAfterEffect}${effectGroup(11, 'clickEffect', 5, 'in', 'cut', 12, 1)}${colorEffect(87, 4, 'fillcolor', 'f59e0b', 'ff0000')}${effectGroup(13, 'withEffect', 7, 'in', 'cut', 14, 1, '0', 0)}${scaleEffect}${rotationEffect}${wipeAnimations}${triggeredEffect}${triggeredMouseOver}${triggeredMouseOut}${triggeredDoubleClick}${elementIterationEffect}${triggeredSlideNext}${triggeredSlidePrevious}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>${speedTriggerSequence}${reverseSpeedSequence}${parentRepeatSequences}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
  const relativeMotion = `<p:animMotion origin="parent"><p:cBhvr><p:cTn id="153" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="92"/></p:tgtEl></p:cBhvr><p:by x="25000" y="0"/></p:animMotion>`
  const fromToMotion = `<p:animMotion origin="layout"><p:cBhvr><p:cTn id="245" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="93"/></p:tgtEl></p:cBhvr><p:from x="25%" y="0%"/><p:to x="50000" y="0"/></p:animMotion>`
  const fromByMotion = `<p:animMotion origin="layout"><p:cBhvr><p:cTn id="246" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="94"/></p:tgtEl></p:cBhvr><p:from x="10000" y="0"/><p:by x="25000" y="0"/></p:animMotion>`
  const toMotion = `<p:animMotion origin="layout"><p:cBhvr><p:cTn id="247" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="95"/></p:tgtEl></p:cBhvr><p:to x="40000" y="0"/></p:animMotion>`
  const rotatedMotionPath = `<p:animMotion origin="layout" pathEditMode="fixed" rAng="5400000" path="M 0.5 0.5 L 0.55 0.5 E"><p:cBhvr><p:cTn id="249" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="97"/></p:tgtEl></p:cBhvr><p:rCtr x="50000" y="50000"/></p:animMotion>`
  const formulaMotionPath = `<p:animMotion origin="layout" pathEditMode="fixed" path="M (#ppt_x #ppt_y) L (#ppt_x+#ppt_w*1.5, #ppt_y+sin(pi/2)*#ppt_h*0.5) E"><p:cBhvr><p:cTn id="250" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="122"/></p:tgtEl></p:cBhvr></p:animMotion>`
  const dynamicFormulaMotionPath = `<p:animMotion origin="layout" pathEditMode="fixed" path="M (#ppt_x-sin(pi*$)/3, #ppt_y) L (#ppt_x, #ppt_y) E"><p:cBhvr><p:cTn id="251" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="123"/></p:tgtEl></p:cBhvr></p:animMotion>`
  const dynamicMotionClick = `<p:par><p:cTn id="278" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>${dynamicFormulaMotionPath}</p:childTnLst></p:cTn></p:par>`
  const dynamicMotionTrigger = triggeredSequence(814, 8, dynamicMotionClick)
  const groupedParentMotion = `<p:animMotion origin="parent" pathEditMode="relative" path="M 0 0 L 0.01 0 E"><p:cBhvr><p:cTn id="252" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="125"/></p:tgtEl></p:cBhvr></p:animMotion>`
  const nestedParentMotion = `<p:animMotion origin="parent" pathEditMode="relative" path="M 0.008202099738 0.007290172061 L 0.018202099738 0.007290172061 E"><p:cBhvr><p:cTn id="253" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="11"/></p:tgtEl></p:cBhvr></p:animMotion>`
  const positionMotions = `${relativeMotion}${fromToMotion}${fromByMotion}${toMotion}${rotatedMotionPath}${formulaMotionPath}${groupedParentMotion}${nestedParentMotion}`
  const fixedMotionPath = `<p:animMotion origin="layout" pathEditMode="fixed" path="M 0 0 L 0.25 0 E"><p:cBhvr><p:cTn id="248" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="96"/></p:tgtEl></p:cBhvr></p:animMotion>`
  const motionEffectWithRelativeMotion = motionEffect.replace('</p:childTnLst>', `${positionMotions}${fixedMotionPath}</p:childTnLst>`)
  const timingWithRelativeMotion = timing.replace(motionEffect, motionEffectWithRelativeMotion).replace('</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>', `${dynamicMotionTrigger}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`)
  assert(timingWithRelativeMotion !== timing, 'The from/to/by motion fixture must be inserted into the main sequence.')
  const wipeShapes = ['right', 'left', 'up', 'down'].map((direction, index) => shape(20 + index, `Wipe ${direction}`, index * 3000000, 4000000, '38bdf8')).join('')
  const animatedShapes = `${shape(2, 'Click fade', 0, 0, 'ef4444')}${shadowedWithPrevious}${shape(4, 'Exit after previous', 6000000, 0, 'f59e0b')}${shape(5, 'Second click', 9000000, 0, '7c3aed')}${shape(6, 'Automatic fade', 0, 3000000, '0891b2')}${textShape(7, 'Paragraph build', 3000000, 3000000)}${plainTextShape(43, 'By letter', 0, 5700000, 'WORD')}${plainTextShape(45, 'By word', 3000000, 5700000, 'ONE TWO')}${shape(44, 'After letters', 6000000, 5700000, '14b8a6')}${shape(46, 'Blinds animation', 9000000, 4000000, '64748b')}${shape(47, 'Checker animation', 3000000, 4000000, '0ea5e9')}${shape(8, 'Scale emphasis', 8000000, 3000000, 'd946ef')}${shape(30, 'Motion path', 1000000, 3000000, '0f766e')}${shape(29, 'Curved motion path', 1000000, 3000000, 'fb923c')}${wipeShapes}${shape(24, 'Wipe exit', 8000000, 0, 'fb7185')}${shape(60, 'Slide entrance and exit', 9300000, 2500000, '22c55e')}${groupedShape}${customGeometryShapeWithGuides}${customGeometryTextShape}${invalidCustomGeometryTextShape}${customGradientShape}${elementIterationGroup}${shape(90, 'Next event target', 9000000, 5500000, '0891b2')}${shape(91, 'Previous event target', 10500000, 5500000, '9333ea')}${shape(97, 'Rotated fixed motion path', 4500000, 3000000, '0ea5e9', 500000, 500000)}${shape(98, 'Fast speed trigger', 8000000, 5200000, '0ea5e9')}${shape(99, 'Fast speed fade', 8500000, 5200000, 'f97316')}${shape(100, 'Reverse speed trigger', 6000000, 5200000, '14b8a6')}${shape(101, 'Reverse speed scale', 6500000, 5200000, 'a855f7')}${shape(102, 'Parent eased fade', 7000000, 5200000, 'facc15')}${shape(103, 'Parent timing after-effect', 7500000, 5200000, '22c55e')}${shape(104, 'Parent speed delayed', 8000000, 5200000, 'fb923c')}${shape(122, 'Formula fixed motion path', 5500000, 3500000, '0d9488')}${shape(123, 'Dynamic formula motion path', 6000000, 3500000, '0d9488')}${parentMotionGroup}`
  const parentRepeatShapes = `${shape(105, 'Parent repeat count trigger', 1000000, 5000000, '0ea5e9')}${shape(106, 'Parent repeat count target', 2000000, 5000000, 'f97316')}${shape(107, 'Parent repeat duration trigger', 3000000, 5000000, '14b8a6')}${shape(108, 'Parent repeat duration target', 4000000, 5000000, 'a855f7')}${shape(109, 'Parent repeat after-effect', 5000000, 5000000, '22c55e')}`
  const parentTimingShapes = `${shape(112, 'Parent reverse speed trigger', 6000000, 5000000, '0ea5e9')}${shape(113, 'Parent reverse speed fade', 7000000, 5000000, 'f97316')}${shape(114, 'Parent auto-reverse trigger', 8000000, 5000000, '14b8a6')}${shape(115, 'Parent auto-reverse fade', 9000000, 5000000, 'a855f7')}${shape(116, 'Parent indefinite repeat trigger', 10000000, 5000000, '14b8a6')}${shape(117, 'Parent indefinite repeat fade', 10500000, 5000000, 'a855f7')}${shape(118, 'Child repeat duration trigger', 2000000, 6200000, '14b8a6')}${shape(119, 'Child repeat duration fade', 4500000, 6200000, 'a855f7')}`
  const normalAutofitShape = plainTextShape(120, 'Normal autofit', 0, 5900000, 'Autofit sample')
    .replace('<a:bodyPr/>', '<a:bodyPr><a:normAutofit fontScale="80%" lnSpcReduction="25%"/></a:bodyPr>')
    .replace('sz="1800"', 'sz="2000"')
    .replace('<a:p><a:r>', '<a:p><a:pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc></a:pPr><a:r>')
  const autofitReferenceShape = plainTextShape(121, 'Autofit reference', 3500000, 5900000, 'Autofit sample').replace('sz="1800"', 'sz="2000"')
  const animatedSlide = `${animatedShapes}${parentRepeatShapes}${parentTimingShapes}${normalAutofitShape}${autofitReferenceShape}${lineShape(42, 'Stroke color', 500000, 6000000, '<a:headEnd type="triangle" w="lg" len="lg"/><a:tailEnd type="diamond" w="sm" len="sm"/>')}${plainTextShape(59, 'Paragraph wipe target', 9000000, 5500000, 'Paragraph wipe target', '<a:ln w="25400"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:ln>')}`
  const animatedSlideWithRelativeMotion = `${animatedSlide}${shape(92, 'Relative position motion', 4500000, 3000000, 'ea580c')}${shape(93, 'From/to motion', 5500000, 3000000, 'c026d3')}${shape(94, 'From/by motion', 6500000, 3000000, 'ca8a04')}${shape(95, 'To-only motion', 7500000, 3000000, '2563eb')}${shape(96, 'Fixed motion path', 4500000, 4000000, '0d9488')}`.replace('<a:ln w="25400">', '<a:ln w="25400"><a:prstDash val="dash"/>')
    zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(animatedSlideWithRelativeMotion)}<p:transition spd="med"><p:zoom dir="in"/></p:transition>${timingWithRelativeMotion}</p:sld>`)
  const lineChartPart = chartPart.replace('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>', '<c:lineChart><c:grouping val="standard"/>').replaceAll('barChart', 'lineChart')
  const pieChartPart = chartPart.replace('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>', '<c:pieChart>').replace('<c:axId val="1"/><c:axId val="2"/>', '').replaceAll('barChart', 'pieChart')
  const doughnutChartPart = pieChartPart.replaceAll('pieChart', 'doughnutChart').replace('</c:doughnutChart>', '<c:firstSliceAng val="90"/><c:holeSize val="40"/></c:doughnutChart>')
  const withValueAxis = (xml, { min, max, majorUnit, numberFormat, reverse = false, categoryPosition = 'b', valuePosition = 'l', categoryCrosses = 'autoZero', valueCrosses = 'autoZero', categoryCrossAt, valueCrossAt, categoryTickMark = 'none', valueTickMark = 'none', categoryTickLabelPosition = 'nextTo', valueTickLabelPosition = 'nextTo', categoryDeleted = false, valueDeleted = false, majorGridlines = true }) => xml.replace('</c:plotArea>', `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${categoryDeleted ? 1 : 0}"/><c:axPos val="${categoryPosition}"/><c:majorTickMark val="${categoryTickMark}"/><c:tickLblPos val="${categoryTickLabelPosition}"/><c:crossAx val="2"/><c:crosses val="${categoryCrosses}"/>${categoryCrossAt === undefined ? '' : `<c:crossesAt val="${categoryCrossAt}"/>`}<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="${reverse ? 'maxMin' : 'minMax'}"/><c:min val="${min}"/><c:max val="${max}"/></c:scaling><c:delete val="${valueDeleted ? 1 : 0}"/><c:axPos val="${valuePosition}"/>${majorGridlines ? '<c:majorGridlines/>' : ''}<c:numFmt formatCode="${numberFormat}" sourceLinked="0"/><c:majorTickMark val="${valueTickMark}"/><c:tickLblPos val="${valueTickLabelPosition}"/><c:minorTickMark val="none"/><c:crossAx val="1"/><c:crosses val="${valueCrosses}"/>${valueCrossAt === undefined ? '' : `<c:crossesAt val="${valueCrossAt}"/>`}<c:crossBetween val="between"/><c:majorUnit val="${majorUnit}"/></c:valAx></c:plotArea>`)
  const withLabelPosition = (xml, position) => xml.replaceAll('<c:dLbls>', `<c:dLbls><c:dLblPos val="${position}"/>`)
  const withPointLabelPosition = (xml, index, position) => xml.replace(`<c:dLbl><c:idx val="${index}"/>`, `<c:dLbl><c:idx val="${index}"/><c:dLblPos val="${position}"/>`)
  const withSeriesPointLabelPosition = (xml, index, position) => xml.replace('</c:val></c:ser>', `</c:val><c:dLbls><c:dLbl><c:idx val="${index}"/><c:dLblPos val="${position}"/></c:dLbl></c:dLbls></c:ser>`)
  const withLegend = (xml, position) => xml.replace('</c:chart>', `<c:legend><c:legendPos val="${position}"/>${position === 'l' ? '<c:layout><c:manualLayout><c:xMode val="edge"/><c:x val="0.1"/><c:yMode val="factor"/><c:y val="0.05"/></c:manualLayout></c:layout>' : ''}</c:legend></c:chart>`)
  const withValueLabels = (xml, chartType, hideFirstSeries = false, hideFirstCategory = false, formatCode = '') => {
    const seriesXml = hideFirstSeries ? xml.replace('</c:val></c:ser>', '</c:val><c:dLbls><c:showVal val="0"/><c:dLbl><c:idx val="1"/><c:numFmt formatCode="0.00" sourceLinked="0"/><c:showVal val="1"/></c:dLbl></c:dLbls></c:ser>') : xml
    const numberFormat = formatCode ? `<c:numFmt formatCode="${formatCode}" sourceLinked="0"/>` : ''
    const labels = `<c:dLbls>${numberFormat}<c:showVal val="1"/>${hideFirstCategory ? '<c:dLbl><c:idx val="0"/><c:showVal val="0"/></c:dLbl>' : ''}</c:dLbls>`
    return chartType === 'pieChart'
      ? seriesXml.replace('</c:pieChart>', `${labels}</c:pieChart>`)
      : seriesXml.replace('<c:axId val="1"/>', `${labels}<c:axId val="1"/>`)
  }
  const pieWithPercentLabels = withValueLabels(pieChartPart, 'pieChart')
    .replace('<c:dLbls><c:showVal val="1"/></c:dLbls>', '<c:dLbls><c:showVal val="0"/><c:showPercent val="1"/></c:dLbls>')
    .replace('</c:val></c:ser>', '</c:val><c:dLbls><c:showCatName val="1"/><c:dLbl><c:idx val="1"/><c:showPercent val="0"/><c:showVal val="1"/><c:separator> / </c:separator></c:dLbl></c:dLbls></c:ser>')
  zip.file('ppt/slides/slide2.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${tree(`${shape(2, 'Second slide', 0, 0, '16a34a')}${table}${gradientShape}${luminanceShape}${groupedShape}${chartFrame(12, 'Quarterly sales chart', 'rIdChart1', 7500000, 4000000, 4300000, 2600000)}${chartFrame(13, 'Quarterly trend chart', 'rIdChart2', 7500000, 100000, 4300000, 2500000)}${chartFrame(14, 'Quarterly share chart', 'rIdChart3', 7500000, 2700000, 4300000, 2500000)}${chartFrame(15, 'Horizontal axis tick chart', 'rIdChart4', 15000000, 8000000, 4300000, 2600000)}${chartFrame(16, 'Category axis deleted chart', 'rIdChart5', 15000000, 11000000, 4300000, 2600000)}${chartFrame(17, 'Value axis deleted chart', 'rIdChart6', 15000000, 14000000, 4300000, 2600000)}`)}<p:transition spd="med"><p:cover dir="ru"/></p:transition></p:sld>`)
  zip.file('ppt/slides/_rels/slide2.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/><Relationship Id="rIdChart3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/><Relationship Id="rIdChart4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart4.xml"/><Relationship Id="rIdChart5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart5.xml"/><Relationship Id="rIdChart6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart6.xml"/></Relationships>`)
  zip.file('ppt/charts/chart1.xml', withLegend(withValueAxis(withPointLabelPosition(withLabelPosition(withValueLabels(chartPart, 'barChart', true, false, '#,##0.0'), 'outEnd'), 1, 'ctr'), { min: 0, max: 30, majorUnit: 10, numberFormat: '#,##0.0', categoryCrosses: 'max', valueCrosses: 'max', categoryTickMark: 'none', valueTickMark: 'out', categoryTickLabelPosition: 'high', valueTickLabelPosition: 'low' }), 'b').replace('</c:legend>', '<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry></c:legend>'))
  zip.file('ppt/charts/chart2.xml', withLegend(withValueAxis(withSeriesPointLabelPosition(withLabelPosition(withValueLabels(lineChartPart, 'lineChart', false, true, '0.0%').replace('<c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/><c:showVal val="1"/>', '<c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/><c:showVal val="1"/><c:showSerName val="1"/><c:separator> | </c:separator>'), 'b'), 2, 'r'), { min: 10, max: 30, majorUnit: 5, numberFormat: '0.0', reverse: true, categoryPosition: 't', valuePosition: 'r', categoryCrossAt: 2, valueCrossAt: 15, categoryTickMark: 'cross', valueTickMark: 'in' }), 'l'))
  zip.file('ppt/charts/chart3.xml', withLegend(pieWithPercentLabels, 'r').replace('</c:legend>', '<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry></c:legend>'))
  const horizontalBarChartPart = chartPart.replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>')
  zip.file('ppt/charts/chart4.xml', withValueAxis(horizontalBarChartPart, { min: 0, max: 30, majorUnit: 10, numberFormat: '0', categoryPosition: 'r', valuePosition: 't', categoryCrosses: 'max', valueCrosses: 'min', categoryTickMark: 'out', valueTickMark: 'in', categoryTickLabelPosition: 'low', valueTickLabelPosition: 'none', majorGridlines: false }))
  zip.file('ppt/charts/chart5.xml', withValueAxis(horizontalBarChartPart, { min: 0, max: 30, majorUnit: 10, numberFormat: '0', categoryPosition: 'r', valuePosition: 't', categoryCrosses: 'max', valueCrosses: 'min', categoryTickMark: 'out', valueTickMark: 'out', categoryDeleted: true }))
  zip.file('ppt/charts/chart6.xml', withLegend(withValueAxis(horizontalBarChartPart, { min: 0, max: 30, majorUnit: 10, numberFormat: '0', categoryPosition: 'r', valuePosition: 't', categoryCrosses: 'max', valueCrosses: 'min', categoryTickMark: 'out', valueTickMark: 'out', valueDeleted: true }), 'r').replace('</c:legend>', '<c:layout><c:manualLayout><c:xMode val="edge"/><c:x val="0.72"/><c:yMode val="factor"/><c:y val="0"/><c:wMode val="edge"/><c:w val="0.95"/><c:hMode val="factor"/><c:h val="0.5"/></c:manualLayout></c:layout><c:overlay val="1"/></c:legend>'))
  zip.file('ppt/charts/chart7.xml', doughnutChartPart)
  const slide2Xml = await zip.file('ppt/slides/slide2.xml').async('string')
  zip.file('ppt/slides/slide2.xml', slide2Xml.replace('</p:spTree>', `${chartFrame(18, 'Quarterly allocation doughnut', 'rIdChart7', 8500000, 5000000, 4200000, 2000000)}</p:spTree>`))
  const slide2Rels = await zip.file('ppt/slides/_rels/slide2.xml.rels').async('string')
  zip.file('ppt/slides/_rels/slide2.xml.rels', slide2Rels.replace('</Relationships>', '<Relationship Id="rIdChart7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart7.xml"/></Relationships>'))
  zip.file('ppt/slides/slide3.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Third slide', 0, 0, '0ea5e9'))}<p:transition spd="med"><p:push dir="r"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide4.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Pull slide', 0, 0, 'f97316'))}<p:transition spd="med"><p:pull dir="ld"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide5.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Split out slide', 0, 0, 'a855f7'))}<p:transition spd="med"><p:split orient="vert" dir="out"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide6.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Split in slide', 0, 0, 'ec4899'))}<p:transition spd="med"><p:split orient="horz" dir="in"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide7.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Circle slide', 0, 0, '14b8a6'))}<p:transition spd="med"><p:circle/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide8.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Diamond slide', 0, 0, 'f43f5e'))}<p:transition spd="med"><p:diamond/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide9.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:future="urn:example:future-transition" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Cut through black slide', 0, 0, '64748b'))}<mc:AlternateContent><mc:Choice Requires="future"><p:transition spd="fast" p14:dur="900"><p:fade/></p:transition></mc:Choice><mc:Fallback><p:transition spd="med"><p:cut thruBlk="1"/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide10.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Cut slide', 0, 0, '64748b'))}<p:transition><p:cut thruBlk="0"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide11.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Honeycomb slide', 0, 0, '0f766e'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="900ms"><p14:honeycomb/></p:transition></mc:Choice><mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide12.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Blinds slide', 0, 0, '0ea5e9'))}<p:transition spd="med"><p:blinds dir="vert"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide13.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Checker slide', 0, 0, '8b5cf6'))}<p:transition spd="med"><p:checker dir="horz"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide14.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Wheel slide', 0, 0, 'f97316'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med"><p14:wheelReverse spokes="8"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="med"><p:wheel spokes="8"/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide15.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${tree(`${mediaPicture(2, 'Embedded video', 500000, 700000, 'videoFile', 'rId2', '<p14:fade in="200ms" out="100ms"/>')}${mediaPicture(3, 'Embedded audio', 6500000, 700000, 'audioFile', 'rId4', '<p14:trim st="100ms" end="100ms"/><p14:fade in="200ms" out="100ms"/>', 'rId3')}${mediaPicture(4, 'Invalid trim', 500000, 4000000, 'audioFile', 'rId4', '<p14:trim st="200ms" end="100ms"/><p14:fade in="200ms" out="100ms"/>')}${mediaPicture(5, 'Cross-slide audio', 6500000, 4000000, 'audioFile', 'rId5', '', 'rId5')}${objectShapeElements}`)}<p:transition spd="med"><p:dissolve/></p:transition>${shapeMaskTiming}</p:sld>`)
  zip.file('ppt/slides/_rels/slide15.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/poster.svg"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/clip.mp4"/><Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/audio.wav"/><Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/audio-short.wav"/><Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/audio-long.wav"/></Relationships>`)
  zip.file('ppt/media/poster.svg', `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 160 90"><rect width="160" height="90" fill="#15231b"/><circle cx="80" cy="45" r="18" fill="#b8f36b"/></svg>`)
  zip.file('ppt/media/clip.mp4', Buffer.from('000000186674797069736F6D0000020069736F6D69736F32', 'hex'))
  zip.file('ppt/media/audio.wav', silentWave())
  zip.file('ppt/media/audio-short.wav', silentWave(0.25))
  zip.file('ppt/media/audio-long.wav', silentWave(2))
   const firstSlideXml = await zip.file('ppt/slides/slide1.xml').async('string')
   const overflowingText = 'Shape autofit must retain readable text and expand its shape when content exceeds the saved bounds. '.repeat(6)
    const dynamicAutoFitText = 'Normal autofit shrinks text to remain inside the saved shape bounds. '.repeat(2)
    const shapeAutoFitFixtures = plainTextShape(122, 'Shape autofit', 0, 5900000, overflowingText).replace('<a:bodyPr/>', '<a:bodyPr><a:spAutoFit/></a:bodyPr>')
     + plainTextShape(123, 'Shape autofit baseline', 3500000, 5900000, overflowingText)
      + plainTextShape(124, 'Normal autofit dynamic', 6500000, 5900000, dynamicAutoFitText).replace('<a:bodyPr/>', '<a:bodyPr><a:normAutofit/></a:bodyPr>')
    const firstSlideWithAutoFit = firstSlideXml.replace('</p:spTree>', `${shapeAutoFitFixtures}</p:spTree>`)
   zip.file('ppt/slides/slide1.xml', firstSlideWithAutoFit
     .replace('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">', 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
     .replace('<p:spTree>', '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBackground"/><a:srcRect l="10000" r="10000"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg><p:spTree>'))
   zip.file('ppt/media/slide-background.png', Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aMd8AAAAASUVORK5CYII=', 'base64'))
   zip.file('ppt/media/tile-background.svg', '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2" viewBox="0 0 2 2"><path d="M0 0h1v1H0z" fill="#ff0000"/><path d="M1 0h1v1H1z" fill="#00ff00"/><path d="M0 1h1v1H0z" fill="#0000ff"/><path d="M1 1h1v1H1z" fill="#ffff00"/></svg>')
   const firstSlideRelsPath = 'ppt/slides/_rels/slide1.xml.rels'
   const firstSlideRels = zip.file(firstSlideRelsPath)
     ? await zip.file(firstSlideRelsPath).async('string')
     : '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
   zip.file(firstSlideRelsPath, firstSlideRels.replace('</Relationships>', '<Relationship Id="rIdBackground" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/slide-background.png"/></Relationships>'))
  const contentTypes = await zip.file('[Content_Types].xml').async('string')
  zip.file('[Content_Types].xml', contentTypes.replace('</Types>', '<Override PartName="/ppt/slides/slide16.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide17.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide18.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide19.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide20.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide21.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide22.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide23.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide24.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide25.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide26.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide27.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide28.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide29.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide30.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide31.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide32.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide33.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide34.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide35.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide36.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide37.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>'))
  const typesWithChart = await zip.file('[Content_Types].xml').async('string')
  zip.file('[Content_Types].xml', typesWithChart.replace('</Types>', '<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/charts/chart4.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/charts/chart5.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/charts/chart6.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'))
  const presentation = await zip.file('ppt/presentation.xml').async('string')
  zip.file('ppt/presentation.xml', presentation.replace('</p:sldIdLst>', '<p:sldId id="271" r:id="rId16"/><p:sldId id="272" r:id="rId17"/><p:sldId id="273" r:id="rId18"/><p:sldId id="274" r:id="rId19"/><p:sldId id="275" r:id="rId20"/><p:sldId id="276" r:id="rId21"/><p:sldId id="277" r:id="rId22"/><p:sldId id="278" r:id="rId23"/><p:sldId id="279" r:id="rId24"/><p:sldId id="280" r:id="rId25"/><p:sldId id="281" r:id="rId26"/><p:sldId id="282" r:id="rId27"/><p:sldId id="283" r:id="rId28"/><p:sldId id="284" r:id="rId29"/><p:sldId id="285" r:id="rId30"/><p:sldId id="286" r:id="rId31"/><p:sldId id="287" r:id="rId32"/><p:sldId id="288" r:id="rId33"/><p:sldId id="289" r:id="rId34"/><p:sldId id="290" r:id="rId35"/><p:sldId id="291" r:id="rId36"/><p:sldId id="292" r:id="rId37"/></p:sldIdLst>'))
  const presentationRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
  zip.file('ppt/_rels/presentation.xml.rels', presentationRels.replace('</Relationships>', '<Relationship Id="rId16" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide16.xml"/><Relationship Id="rId17" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide17.xml"/><Relationship Id="rId18" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide18.xml"/><Relationship Id="rId19" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide19.xml"/><Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide20.xml"/><Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide21.xml"/><Relationship Id="rId22" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide22.xml"/><Relationship Id="rId23" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide23.xml"/><Relationship Id="rId24" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide24.xml"/><Relationship Id="rId25" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide25.xml"/><Relationship Id="rId26" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide26.xml"/><Relationship Id="rId27" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide27.xml"/><Relationship Id="rId28" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide28.xml"/><Relationship Id="rId29" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide29.xml"/><Relationship Id="rId30" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide30.xml"/><Relationship Id="rId31" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide31.xml"/><Relationship Id="rId32" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide32.xml"/><Relationship Id="rId33" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide33.xml"/><Relationship Id="rId34" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide34.xml"/><Relationship Id="rId35" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide35.xml"/><Relationship Id="rId36" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide36.xml"/><Relationship Id="rId37" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide37.xml"/></Relationships>'))
  zip.file('ppt/slides/slide16.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Random bars slide', 0, 0, '14b8a6'))}<p:transition spd="med"><p:randomBar dir="horz"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide17.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId3"/><a:tile sx="50000" sy="200000" algn="br" flip="xy" tx="914400" ty="-457200"/></a:blipFill></p:bgPr></p:bg>${tree(`${shape(2, 'Vertical random bars slide', 0, 0, '0ea5e9')}${placeholderShape}`)}<p:transition spd="med"><p:randomBar dir="vert"/></p:transition></p:sld>`)
  const smartArtFrame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="90" name="SmartArt text fallback fixture"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="9200000" y="5400000"/><a:ext cx="2500000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId4"/></a:graphicData></a:graphic></p:graphicFrame>`
  const oleFrame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="92" name="OLE preview fixture"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="9200000" y="4200000"/><a:ext cx="2500000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Quarterly workbook" progId="Excel.Sheet.12" r:id="rId5"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="93" name="OLE preview image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2500000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`
  const oleNoPreviewFrame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="94" name="OLE no-preview fixture"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="6500000" y="5400000"/><a:ext cx="2200000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Linked report" progId="Word.Document.12" r:id="rId5"><p:link/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`
  const smartArtSlide17Xml = await zip.file('ppt/slides/slide17.xml').async('string')
  zip.file('ppt/slides/slide17.xml', smartArtSlide17Xml.replace('</p:spTree>', `${smartArtFrame}${oleFrame}${oleNoPreviewFrame}</p:spTree>`))
  zip.file('ppt/slides/slide18.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Staggered strips slide', 0, 0, 'f59e0b'))}<p:transition spd="med"><p:strips dir="ru"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide19.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Plus transition slide', 0, 0, '8b5cf6'))}<p:transition spd="med"><p:plus/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide20.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Wedge transition slide', 0, 0, 'ec4899'))}<p:transition spd="med"><p:wedge/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide21.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Newsflash transition slide', 0, 0, 'f97316'))}<p:transition spd="med"><p:newsflash/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide22.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Random transition slide', 0, 0, '22c55e'))}<p:transition spd="med"><p:random/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide23.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(shape(2, 'Fade through black slide', 0, 0, '3b82f6'))}<p:transition spd="slow"><p:fade thruBlk="1"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide24.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${tree(shape(2, 'Long explicit transition', 0, 0, 'f59e0b'))}<p:transition advClick="0" advTm="5000" p14:dur="45000"><p:comb dir="vert"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide25.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${tree(`${shape(2, 'Fast transition', 0, 0, '10b981')}${autoFitText(3, 'Auto fit calibration', 3000000, 3000000, 6000000, 583565, 3200, 'Office slide calibration')}${autoFitText(4, 'Small auto-fit calibration', 0, 4000000, 3000000, 200000, 1000, 'Small text')}`)}<p:transition spd="fast"><p:comb dir="horz"/></p:transition></p:sld>`)
  zip.file('ppt/slides/slide26.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Flash transition', 0, 0, 'dc2626'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="900ms"><p14:flash/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide27.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Doors transition', 0, 0, 'ca8a04'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="600ms"><p14:doors dir="vert"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide28.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Window transition', 0, 0, '0891b2'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="700ms"><p14:window dir="horz"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide29.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Prism transition', 0, 0, '7c3aed'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="850ms"><p14:prism dir="r" isInverted="1" isContent="1"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide30.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Flip transition', 0, 0, '0891b2'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="800ms"><p14:flip dir="r"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide31.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Reveal transition', 0, 0, '0f766e'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="1050ms"><p14:reveal dir="l" thruBlk="true"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide32.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Switch transition', 0, 0, '6d28d9'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="950ms"><p14:switch dir="r"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide33.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Pan transition', 0, 0, '6d28d9'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="900ms"><p14:pan dir="ld"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide34.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Ripple transition', 0, 0, '0284c7'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="1200ms"><p14:ripple dir="ld"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide35.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Glitter diamond left', 0, 0, 'd946ef'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="1000ms"><p14:glitter dir="l" pattern="diamond"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide36.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Glitter diamond right down', 0, 0, 'a855f7'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="1100ms"><p14:glitter dir="rd" pattern="diamond"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide37.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Glitter hexagon down', 0, 0, '0ea5e9'))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="1200ms"><p14:glitter dir="d" pattern="hexagon"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  const shredSlide = (title, color, attributes) => `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, title, 0, 0, color))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="1000ms"><p14:shred ${attributes}/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`
  zip.file('ppt/slides/slide38.xml', shredSlide('Shred default strip in', 'eab308', ''))
  zip.file('ppt/slides/slide39.xml', shredSlide('Shred rectangle out', '8b5cf6', 'dir="out" pattern="rectangle"'))
  const flythroughSlide = (title, color, attributes) => `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, title, 0, 0, color))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="900ms"><p14:flythrough ${attributes}/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`
  zip.file('ppt/slides/slide40.xml', flythroughSlide('Flythrough default in', 'f97316', ''))
  zip.file('ppt/slides/slide41.xml', flythroughSlide('Flythrough bounced out', '2563eb', 'dir="out" hasBounce="true"'))
  const warpSlide = (title, color, attributes) => `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, title, 0, 0, color))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="850ms"><p14:warp ${attributes}/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`
  zip.file('ppt/slides/slide42.xml', warpSlide('Warp default out', '14b8a6', ''))
  zip.file('ppt/slides/slide43.xml', warpSlide('Warp explicit in', 'ec4899', 'dir="in"'))
  const vortexSlide = (title, color, attributes) => `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, title, 0, 0, color))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="850ms"><p14:vortex ${attributes}/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`
  zip.file('ppt/slides/slide44.xml', vortexSlide('Vortex default left', '14b8a6', ''))
  zip.file('ppt/slides/slide45.xml', vortexSlide('Vortex right', 'ec4899', 'dir="r"'))
  const directionTransitionSlide = (effect, title, color, attributes) => `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, title, 0, 0, color))}<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="med" p14:dur="850ms"><p14:${effect} ${attributes}/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`
  zip.file('ppt/slides/slide46.xml', directionTransitionSlide('ferris', 'Ferris default left', 'eab308', ''))
  zip.file('ppt/slides/slide47.xml', directionTransitionSlide('ferris', 'Ferris right', '8b5cf6', 'dir="r"'))
  zip.file('ppt/slides/slide48.xml', directionTransitionSlide('gallery', 'Gallery default left', '06b6d4', ''))
  zip.file('ppt/slides/slide49.xml', directionTransitionSlide('gallery', 'Gallery right', 'f97316', 'dir="r"'))
  zip.file('ppt/slides/slide50.xml', directionTransitionSlide('conveyor', 'Conveyor default left', '10b981', ''))
  zip.file('ppt/slides/slide51.xml', directionTransitionSlide('conveyor', 'Conveyor right', 'd946ef', 'dir="r"'))
  zip.file('ppt/slides/slide52.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Conveyor right', 4200000, 2200000, 'f59e0b', 3400000, 1700000))}<mc:AlternateContent><mc:Choice Requires="p159"><p:transition spd="med" p14:dur="1000ms"><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  zip.file('ppt/slides/slide53.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Conveyor right', 6000000, 800000, '8b5cf6', 2400000, 2000000))}<mc:AlternateContent><mc:Choice Requires="p159"><p:transition spd="med" p14:dur="650ms"><p159:morph option="byWord"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  const presetNames = ['fallOver', 'drape', 'curtains', 'wind', 'prestige', 'fracture', 'crush', 'peelOff', 'pageCurlDouble', 'pageCurlSingle', 'airplane', 'origami']
  presetNames.forEach((presetName, index) => {
    const attributes = presetName === 'fallOver' ? 'invY="1"' : presetName === 'peelOff' ? 'invX="1"' : ''
    zip.file(`ppt/slides/slide${54 + index}.xml`, `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, `Preset ${presetName}`, 0, 0, '0ea5e9'))}<mc:AlternateContent><mc:Choice Requires="p15"><p:transition spd="med" p14:dur="850ms"><p15:prstTrans prst="${presetName}" ${attributes}/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  })
  zip.file('ppt/slides/slide66.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${tree(shape(2, 'Preset invalid fallback', 0, 0, 'f97316'))}<mc:AlternateContent><mc:Choice Requires="p15"><p:transition spd="med" p14:dur="850ms"><p15:prstTrans prst="notARealPreset"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`)
  const visibilityTiming = `<p:timing><p:tnLst><p:par><p:cTn id="12028" dur="indefinite" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek" prevAc="skipTimed"><p:cTn id="12029" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="12030" dur="indefinite" nodeType="clickEffect"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>${visibilityAnimation}${visibilityKeyframes}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
  const visibilitySlide = await zip.file('ppt/slides/slide66.xml').async('string')
  zip.file('ppt/slides/slide66.xml', visibilitySlide
    .replace('</p:spTree>', `${shape(130, 'Visibility from-to', 0, 5000000, 'c084fc', 100000, 100000)}${shape(131, 'Visibility keyframes', 200000, 5000000, 'c084fc', 100000, 100000)}</p:spTree>`)
    .replace('</p:sld>', `${visibilityTiming}</p:sld>`))
  const textScale = `<p:animScale><p:cBhvr><p:cTn id="671" dur="800" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr><p:by x="160000" y="125000"/></p:animScale>`
  const textRotation = `<p:animRot by="1800000"><p:cBhvr><p:cTn id="672" dur="800" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="1" end="3"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot><p:animMotion origin="layout" path="M 0 0 L 0.12 0.05 E" pathEditMode="relative"><p:cBhvr><p:cTn id="676" dur="800" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="0" end="1"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`
  const nestedPresetScale = `<p:par><p:cTn id="681" dur="indefinite" nodeType="withEffect" presetClass="emph" presetID="6" presetSubtype="0"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animScale><p:cBhvr><p:cTn id="682" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="2" end="3"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr><p:by x="130000" y="130000"/></p:animScale></p:childTnLst></p:cTn></p:par>`
  const characterBlinds = `<p:par><p:cTn id="683" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="blinds(horizontal)"><p:cBhvr><p:cTn id="684" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="1" end="3"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const paragraphBlinds = `<p:par><p:cTn id="685" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="blinds(vertical)"><p:cBhvr><p:cTn id="686" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="687"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const paragraphChecker = `<p:par><p:cTn id="689" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="checkerboard(across)"><p:cBhvr><p:cTn id="690" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="688"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const paragraphBarn = `<p:par><p:cTn id="693" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="barn(inHorizontal)"><p:cBhvr><p:cTn id="694" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="689"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const characterBarn = `<p:par><p:cTn id="695" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="barn(inVertical)"><p:cBhvr><p:cTn id="696" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="1" end="3"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const paragraphRandomBars = `<p:par><p:cTn id="699" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="randombar(vertical)"><p:cBhvr><p:cTn id="700" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="690"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const paragraphStrips = `<p:par><p:cTn id="701" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="strips(upLeft)"><p:cBhvr><p:cTn id="702" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="691"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const characterStrips = `<p:par><p:cTn id="703" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="strips(downRight)"><p:cBhvr><p:cTn id="704" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="692"><p:txEl><p:charRg st="1" end="3"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const paragraphDissolve = `<p:par><p:cTn id="705" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="dissolve"><p:cBhvr><p:cTn id="706" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="693"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const characterDissolve = `<p:par><p:cTn id="707" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="dissolve"><p:cBhvr><p:cTn id="708" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="694"><p:txEl><p:charRg st="1" end="3"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const paragraphWheel = `<p:par><p:cTn id="709" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="wheel(4)"><p:cBhvr><p:cTn id="710" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="695"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const characterWedge = `<p:par><p:cTn id="711" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="wedge"><p:cBhvr><p:cTn id="712" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="696"><p:txEl><p:charRg st="1" end="3"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const paragraphShapeMask = `<p:par><p:cTn id="713" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="diamond(in)"><p:cBhvr><p:cTn id="714" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="697"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  zip.file('ppt/slides/slide67.xml', `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld>${tree(plainTextShape(670, 'Character transform target', 2500000, 1800000, 'PPTX') + plainTextShape(687, 'Paragraph blinds target', 2500000, 3200000, 'Blinds') + plainTextShape(688, 'Paragraph checker target', 5500000, 3200000, 'Checker') + plainTextShape(689, 'Paragraph barn target', 8500000, 3200000, 'Barn') + plainTextShape(690, 'Paragraph random bars target', 10500000, 3200000, 'Random bars') + plainTextShape(691, 'Paragraph strips target', 8500000, 4000000, 'Strips') + plainTextShape(692, 'Character strips target', 10500000, 4000000, 'ABCD') + plainTextShape(693, 'Paragraph dissolve target', 0, 4800000, 'Dissolve') + plainTextShape(694, 'Character dissolve target', 3000000, 4800000, 'ABCD') + plainTextShape(695, 'Paragraph wheel target', 6000000, 4800000, 'Wheel') + plainTextShape(696, 'Character wedge target', 9000000, 4800000, 'ABCD') + plainTextShape(697, 'Paragraph shape target', 12000000, 4800000, 'Shape'))}</p:cSld><p:timing><p:tnLst><p:par><p:cTn id="673" dur="indefinite" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1"><p:cTn id="674" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="675" dur="indefinite" nodeType="clickEffect"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>${textScale}${textRotation}${nestedPresetScale}${characterBlinds}${paragraphBlinds}${paragraphChecker}${paragraphBarn}${characterBarn}${paragraphRandomBars}${paragraphStrips}${characterStrips}${paragraphDissolve}${characterDissolve}${paragraphWheel}${characterWedge}${paragraphShapeMask}${characterRangeOpacity}${setCharacterRangeOpacity}${characterRangeVisibility}${paragraphRangeVisibility}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>`)
  const withGroupTextColor = `<p:par><p:cTn id="677" dur="indefinite" nodeType="withGroup"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animClr><p:cBhvr><p:cTn id="678" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="3" end="4"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="000000"/></p:from><p:to><a:srgbClr val="ff0000"/></p:to></p:animClr></p:childTnLst></p:cTn></p:par>`
  const afterGroupTextColor = `<p:par><p:cTn id="679" dur="indefinite" nodeType="afterGroup"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animClr><p:cBhvr><p:cTn id="680" dur="300" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="3" end="4"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="ff0000"/></p:from><p:to><a:srgbClr val="0000ff"/></p:to></p:animClr></p:childTnLst></p:cTn></p:par>`
  const characterChecker = `<p:par><p:cTn id="691" dur="indefinite" nodeType="withEffect" presetClass="entr" presetID="22"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="checkerboard(across)"><p:cBhvr><p:cTn id="692" dur="500" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="670"><p:txEl><p:charRg st="1" end="3"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`
  const textRangeSlideXml = await zip.file('ppt/slides/slide67.xml').async('string')
  zip.file('ppt/slides/slide67.xml', textRangeSlideXml.replace('</p:childTnLst></p:cTn></p:seq>', `${withGroupTextColor}${afterGroupTextColor}${characterChecker}</p:childTnLst></p:cTn></p:seq>`))
  const morphContentTypes = await zip.file('[Content_Types].xml').async('string')
  zip.file('[Content_Types].xml', morphContentTypes.replace('</Types>', '<Override PartName="/ppt/slides/slide52.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide53.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'))
  const presetContentTypes = await zip.file('[Content_Types].xml').async('string')
  zip.file('[Content_Types].xml', presetContentTypes.replace('</Types>', Array.from({ length: 14 }, (_, index) => `<Override PartName="/ppt/slides/slide${54 + index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') + '</Types>'))
  const morphPresentationRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
  zip.file('ppt/_rels/presentation.xml.rels', morphPresentationRels.replace('</Relationships>', '<Relationship Id="rId52" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide52.xml"/><Relationship Id="rId53" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide53.xml"/></Relationships>'))
  const presetPresentationRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
  zip.file('ppt/_rels/presentation.xml.rels', presetPresentationRels.replace('</Relationships>', Array.from({ length: 14 }, (_, index) => `<Relationship Id="rId${54 + index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${54 + index}.xml"/>`).join('') + '</Relationships>'))
  zip.file('ppt/slides/_rels/slide17.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/tile-background.svg"/></Relationships>`)
  const slide17Rels = await zip.file('ppt/slides/_rels/slide17.xml.rels').async('string')
  zip.file('ppt/slides/_rels/slide17.xml.rels', slide17Rels.replace('</Relationships>', '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data2.xml"/></Relationships>'))
  zip.file('ppt/diagrams/data1.xml', `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{00}" type="doc"/><dgm:pt modelId="{01}" type="node"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>First diagram node</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="{02}" type="node"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Second diagram node</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="{10}" srcId="{00}" destId="{01}" srcOrd="0" destOrd="0" type="parOf"/><dgm:cxn modelId="{11}" srcId="{01}" destId="{02}" srcOrd="0" destOrd="0" type="parOf"/></dgm:cxnLst><dgm:whole/><dgm:extLst/></dgm:dataModel>`)
  zip.file('ppt/diagrams/data2.xml', `<?xml version="1.0"?><dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{10}" type="doc"/><dgm:pt modelId="{11}" type="node"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Fallback root</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="{12}" type="node"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Fallback child</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="{13}" srcId="{10}" destId="{11}" srcOrd="0" destOrd="0" type="parOf"/><dgm:cxn modelId="{14}" srcId="{11}" destId="{12}" srcOrd="0" destOrd="0" type="parOf"/></dgm:cxnLst><dgm:whole/><dgm:extLst/></dgm:dataModel>`)
  zip.file('ppt/diagrams/drawing1.xml', `<?xml version="1.0"?><dgm:drawing xmlns:dgm="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:spTree><dgm:nvGrpSpPr><dgm:cNvPr id="0" name=""/><dgm:cNvGrpSpPr/><dgm:nvPr/></dgm:nvGrpSpPr><dgm:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="400000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="400000"/></a:xfrm></dgm:grpSpPr><dgm:sp><dgm:nvSpPr><dgm:cNvPr id="101" name="Root node"/><dgm:cNvSpPr/><dgm:nvPr/></dgm:nvSpPr><dgm:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="450000" cy="150000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0f766e"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="134e4a"/></a:solidFill></a:ln></dgm:spPr><dgm:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1400"/><a:t>Cached diagram root</a:t></a:r></a:p></dgm:txBody></dgm:sp><dgm:sp><dgm:nvSpPr><dgm:cNvPr id="102" name="Child node"/><dgm:cNvSpPr/><dgm:nvPr/></dgm:nvSpPr><dgm:spPr><a:xfrm><a:off x="500000" y="250000"/><a:ext cx="450000" cy="150000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="f59e0b"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="b45309"/></a:solidFill></a:ln></dgm:spPr><dgm:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1400" b="1"/><a:t>Cached diagram child</a:t></a:r></a:p></dgm:txBody></dgm:sp></dgm:spTree></dgm:drawing>`)
  const contentTypesWithOle = await zip.file('[Content_Types].xml').async('string')
  zip.file('[Content_Types].xml', contentTypesWithOle.replace('</Types>', '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>'))
  zip.file('ppt/embeddings/oleObject1.bin', Buffer.from('Fixture OLE bytes'))
  zip.file('ppt/notesSlides/notesSlide1.xml', `<?xml version="1.0"?><p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Introduce the main point.</a:t></a:r><a:endParaRPr lang="en-US"/></a:p><a:p><a:r><a:rPr lang="en-US"/><a:t>Pause for questions.</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`)
  zip.file('ppt/notesSlides/_rels/notesSlide1.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide17.xml"/></Relationships>`)
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0"?><p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" type="obj">${tree(`<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="900000" y="900000"/><a:ext cx="9000000" cy="4500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${shadow}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl3pPr algn="ctr"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:buChar char="→"/><a:defRPr sz="3200"/></a:lvl3pPr></a:lstStyle><a:p><a:pPr lvl="2"/><a:endParaRPr i="1"/></a:p></p:txBody></p:sp>`)}</p:sldLayout>`)
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`)
  zip.file('ppt/slideMasters/slideMaster1.xml', `<?xml version="1.0"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:txStyles><p:titleStyle/><p:bodyStyle><a:lvl1pPr algn="r" marL="360000" indent="-180000"><a:spcBef><a:spcPts val="1200"/></a:spcBef></a:lvl1pPr><a:lvl3pPr algn="r" marL="360000" indent="-180000"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:buChar char="•"/><a:defRPr sz="2800" b="1" strike="sngStrike"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:defRPr></a:lvl3pPr></p:bodyStyle><p:otherStyle/></p:txStyles></p:sldMaster>`)
  const slide15Xml = await zip.file('ppt/slides/slide15.xml').async('string')
  zip.file('ppt/slides/slide15.xml', slide15Xml.replace('</p:spTree>', `${objectStripsElements}${objectBarnElements}${plainTextShape(73, 'Character range wipe target', 9300000, 5500000, 'WIPE')}</p:spTree>`))
  const shredContentTypes = await zip.file('[Content_Types].xml').async('string')
  zip.file('[Content_Types].xml', shredContentTypes.replace('</Types>', '<Override PartName="/ppt/slides/slide38.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide39.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide40.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide41.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide42.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide43.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide44.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide45.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide46.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide47.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide48.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide49.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide50.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide51.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'))
  const shredPresentation = await zip.file('ppt/presentation.xml').async('string')
  zip.file('ppt/presentation.xml', shredPresentation.replace('</p:sldIdLst>', '<p:sldId id="301" r:id="rId38"/><p:sldId id="302" r:id="rId39"/><p:sldId id="303" r:id="rId40"/><p:sldId id="304" r:id="rId41"/><p:sldId id="305" r:id="rId42"/><p:sldId id="306" r:id="rId43"/><p:sldId id="307" r:id="rId44"/><p:sldId id="308" r:id="rId45"/><p:sldId id="309" r:id="rId46"/><p:sldId id="310" r:id="rId47"/><p:sldId id="311" r:id="rId48"/><p:sldId id="312" r:id="rId49"/><p:sldId id="313" r:id="rId50"/><p:sldId id="314" r:id="rId51"/></p:sldIdLst>'))
  const orderedMorphPresentation = await zip.file('ppt/presentation.xml').async('string')
  zip.file('ppt/presentation.xml', orderedMorphPresentation.replace('r:id="rId51"/></p:sldIdLst>', 'r:id="rId51"/><p:sldId id="315" r:id="rId52"/><p:sldId id="316" r:id="rId53"/></p:sldIdLst>'))
  const orderedPresetPresentation = await zip.file('ppt/presentation.xml').async('string')
  zip.file('ppt/presentation.xml', orderedPresetPresentation.replace('r:id="rId53"/></p:sldIdLst>', 'r:id="rId53"/>' + Array.from({ length: 14 }, (_, index) => `<p:sldId id="${317 + index}" r:id="rId${54 + index}"/>`).join('') + '</p:sldIdLst>'))
  const shredPresentationRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
  zip.file('ppt/_rels/presentation.xml.rels', shredPresentationRels.replace('</Relationships>', '<Relationship Id="rId38" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide38.xml"/><Relationship Id="rId39" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide39.xml"/><Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide40.xml"/><Relationship Id="rId41" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide41.xml"/><Relationship Id="rId42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide42.xml"/><Relationship Id="rId43" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide43.xml"/><Relationship Id="rId44" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide44.xml"/><Relationship Id="rId45" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide45.xml"/><Relationship Id="rId46" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide46.xml"/><Relationship Id="rId47" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide47.xml"/><Relationship Id="rId48" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide48.xml"/><Relationship Id="rId49" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide49.xml"/><Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide50.xml"/><Relationship Id="rId51" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide51.xml"/></Relationships>'))
  const slide17Xml = await zip.file('ppt/slides/slide17.xml').async('string')
  const wordArtVerticalRtl = plainTextShape(79, 'WordArt vertical RTL', 3500000, 3600000, 'This is')
    .replace('cy="800000"', 'cy="2500000"')
    .replace('</a:p>', '</a:p><a:p><a:r><a:rPr sz="1800"/><a:t>some text.</a:t></a:r></a:p>')
    .replace('<a:bodyPr/>', '<a:bodyPr vert="wordArtVertRtl"/>')
  const verticalOrientationFixtures = plainTextShape(74, 'Vertical 270 text', 6000000, 5800000, 'Vertical 270').replace('<a:bodyPr/>', '<a:bodyPr vert="vert270"/>')
    + plainTextShape(75, 'WordArt vertical', 9000000, 5800000, 'WordArt').replace('<a:bodyPr/>', '<a:bodyPr vert="wordArtVert"/>')
    + plainTextShape(76, 'East Asian vertical text', 6000000, 4700000, '東西南北').replace('<a:bodyPr/>', '<a:bodyPr vert="eaVert"/>')
    + plainTextShape(77, 'Explicit horizontal text', 9000000, 4700000, 'Horizontal').replace('<a:bodyPr/>', '<a:bodyPr vert="horz"/>')
    + plainTextShape(78, 'Mongolian vertical', 0, 5000000, 'ᠮᠣᠩᠭᠣᠯ').replace('<a:bodyPr/>', '<a:bodyPr vert="mongolianVert"/>')
    + wordArtVerticalRtl
    + plainTextShape(80, 'Two-column text', 0, 3000000, 'First column text keeps flowing until its height is full, then the remaining words continue in the second column with the configured gap.').replace('<a:bodyPr/>', '<a:bodyPr numCol="2" spcCol="457200" rtlCol="0"/>')
    + plainTextShape(81, 'Two-column RTL text', 3000000, 3000000, 'The rightmost text column comes first, then overflow continues in the left column while paragraph text remains left-to-right.').replace('<a:bodyPr/>', '<a:bodyPr numCol="2" spcCol="457200" rtlCol="1"/>')
    + plainTextShape(82, 'RTL paragraph direction', 6000000, 3000000, 'Paragraph rtl controls this paragraph direction independently from the text column order.').replace('<a:p>', '<a:p><a:pPr rtl="1"/>').replace('<a:bodyPr/>', '<a:bodyPr numCol="2" spcCol="457200" rtlCol="1"/>')
  zip.file('ppt/slides/slide17.xml', slide17Xml.replace('</p:spTree>', verticalOrientationFixtures + '</p:spTree>'))
  const layoutPath = 'ppt/slideLayouts/slideLayout1.xml'
  const layoutXml = await zip.file(layoutPath).async('string')
  const layoutWithStyles = layoutXml.replace('<a:noFill/>', '<a:noFill/><a:ln w="25400"><a:solidFill><a:srgbClr val="ef4444"/></a:solidFill><a:prstDash val="dash"/></a:ln>').replace('<a:bodyPr/>', '<a:bodyPr wrap="none" vert="vert"/>')
  assert(layoutWithStyles !== layoutXml && layoutWithStyles.includes('<a:bodyPr wrap="none" vert="vert"/>'), 'The layout placeholder line, no-wrap, and vertical-text fixtures must be inserted.')
  zip.file(layoutPath, layoutWithStyles)
  const imageContentTypes = await zip.file('[Content_Types].xml').async('string')
  zip.file('[Content_Types].xml', imageContentTypes.replace('</Types>', '<Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/></Types>'))
  await writeFile(fixturePath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

async function connectToPage() {
  let targets
  for (let attempt = 0; attempt < 60; attempt++) {
    try {
      targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json()
      const page = targets.find(target => target.type === 'page' && target.url.startsWith(baseUrl))
      if (page?.webSocketDebuggerUrl) return page.webSocketDebuggerUrl
    } catch {}
    await delay(100)
  }
  throw new Error('The headless browser did not expose the player page.')
}

try {
  if (!inputPptx) await createFixture()
  browser = spawn(browserPath, [
    '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--autoplay-policy=no-user-gesture-required',
    '--remote-allow-origins=*', `--remote-debugging-port=${debugPort}`,
    `--user-data-dir=${profilePath}`, baseUrl,
  ], { stdio: 'ignore', windowsHide: true })

  socket = new WebSocket(await connectToPage())
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })

  let id = 0
  const pending = new Map()
  const runtimeErrors = []
  socket.addEventListener('message', event => {
    const message = JSON.parse(event.data)
    if (message.method === 'Runtime.exceptionThrown') runtimeErrors.push(message.params.exceptionDetails.text)
    if (!message.id || !pending.has(message.id)) return
    const request = pending.get(message.id)
    pending.delete(message.id)
    message.error ? request.reject(new Error(message.error.message)) : request.resolve(message.result)
  })
  const call = (method, params = {}) => new Promise((resolve, reject) => {
    const requestId = ++id
    pending.set(requestId, { resolve, reject })
    socket.send(JSON.stringify({ id: requestId, method, params }))
  })
  closeBrowser = async () => { await call('Browser.close').catch(() => {}) }
  const evaluate = async expression => {
    const result = await call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text)
    return result.result.value
  }
  const state = () => evaluate(`(() => {
    const currentHost = document.querySelector('.stage-shell > .slide-host')
    const underlay = document.querySelector('.slide-transition-underlay')
    const outgoingHost = document.querySelector('.slide-transition-underlay > .slide-host')
    const hostRect = currentHost?.getBoundingClientRect()
    const outgoingRect = outgoingHost?.getBoundingClientRect()
    const frame = currentHost?.querySelector('.slide-frame')
    return ({
    now: performance.now(),
    title: document.querySelector('.deck-title')?.textContent,
    slide: document.querySelector('.slide-position')?.textContent?.replace(/\\s+/g, ' ').trim(),
    transitionEffect: currentHost?.dataset.transitionEffect,
    transitionPreset: currentHost?.dataset.transitionPreset,
    slideBackground: frame && { image: getComputedStyle(frame).backgroundImage, size: getComputedStyle(frame).backgroundSize, position: getComputedStyle(frame).backgroundPosition, repeat: getComputedStyle(frame).backgroundRepeat, frameWidth: frame.getBoundingClientRect().width },
    hostFrame: hostRect && { left: hostRect.left, top: hostRect.top, width: hostRect.width, height: hostRect.height },
    shapes: [...(currentHost?.querySelectorAll('.slide-element') || [])].map(element => {
      const rect = element.getBoundingClientRect()
      const frame = element.closest('.slide-frame').getBoundingClientRect()
      const toneFilter = element.querySelector('.image-tone-filter')
      return {
        name: element.title,
        kind: element.dataset.elementKind,
        imageSource: element.querySelector('.slide-image')?.getAttribute('src'),
        imageFilter: element.querySelector('.slide-image') ? getComputedStyle(element.querySelector('.slide-image')).filter : undefined,
        imageCropStyle: element.querySelector('.slide-image') ? Object.fromEntries(['width', 'height', 'left', 'top'].map(property => [property, element.querySelector('.slide-image').style[property]])) : undefined,
        imageTone: toneFilter ? {
          gain: {slope: toneFilter.querySelector('.image-tone-gain')?.getAttribute('slope')},
          blacklevel: {intercept: toneFilter.querySelector('.image-tone-blacklevel')?.getAttribute('intercept')},
          gamma: {exponent: toneFilter.querySelector('.image-tone-gamma')?.getAttribute('exponent')},
        } : undefined,
        visibility: getComputedStyle(element).visibility,
        opacity: getComputedStyle(element).opacity,
        backgroundImage: getComputedStyle(element).backgroundImage,
        maskImage: getComputedStyle(element).maskImage,
        clipPath: getComputedStyle(element).clipPath,
        transform: getComputedStyle(element).transform,
        wipeProgress: getComputedStyle(element).getPropertyValue('--wipe-progress'),
        blindsStripeSize: getComputedStyle(element).getPropertyValue('--blinds-stripe-size').trim(),
        boxShadow: getComputedStyle(element).boxShadow,
        backgroundColor: getComputedStyle(element).backgroundColor,
        borderColor: getComputedStyle(element).borderColor,
        borderStyle: getComputedStyle(element).borderStyle,
        borderWidth: getComputedStyle(element).borderWidth,
        lineColor: element.querySelector('.line-render line') ? getComputedStyle(element.querySelector('.line-render line')).stroke : undefined,
        line: (() => { const line = element.querySelector('.line-render line'); return line ? { x1: Number(line.getAttribute('x1')), y1: Number(line.getAttribute('y1')), x2: Number(line.getAttribute('x2')), y2: Number(line.getAttribute('y2')), markerStart: line.getAttribute('marker-start'), markerEnd: line.getAttribute('marker-end'), headPath: element.querySelector('.line-render marker[id$="-head"] path')?.getAttribute('d'), tailPath: element.querySelector('.line-render marker[id$="-tail"] path')?.getAttribute('d') } : undefined })(),
        customPaths: [...element.querySelectorAll('.custom-geometry path')].map(path => ({ d: path.getAttribute('d'), fill: getComputedStyle(path).fill, stroke: getComputedStyle(path).stroke, strokeWidth: getComputedStyle(path).strokeWidth })),
        frameRect: { left: rect.left - frame.left, top: rect.top - frame.top, width: rect.width, height: rect.height, frameWidth: frame.width, frameHeight: frame.height },
      }
    }),
    media: {
      videos: [...(currentHost?.querySelectorAll('.slide-video') || [])].map(element => ({ name: element.closest('.slide-element')?.title, src: element.getAttribute('src'), poster: element.poster, controls: element.controls, autoplay: element.autoplay, currentTime: element.currentTime, duration: element.duration, volume: element.volume, muted: element.muted, opacity: getComputedStyle(element).opacity })),
      audios: [...(currentHost?.querySelectorAll('.slide-audio') || [])].map(element => ({ id: element.dataset.mediaId, name: element.closest('.slide-element')?.title, src: element.getAttribute('src'), controls: element.controls, autoplay: element.autoplay, duration: element.duration, currentTime: element.currentTime, volume: element.volume, muted: element.muted, readyState: element.readyState })),
      warnings: [...(currentHost?.querySelectorAll('.media-warning') || [])].map(element => ({ name: element.closest('.slide-element')?.title, message: element.textContent })),
    },
    paragraphs: [...(currentHost?.querySelectorAll('.text-paragraph') || [])].map(element => ({ text: element.textContent, visibility: getComputedStyle(element).visibility, opacity: getComputedStyle(element).opacity, maskImage: getComputedStyle(element).maskImage, wipeProgress: getComputedStyle(element).getPropertyValue('--wipe-progress').trim(), color: getComputedStyle(element.querySelector('.text-run') || element).color, textAlign: getComputedStyle(element).textAlign, lineHeight: element.style.lineHeight, marginLeft: element.style.marginLeft, marker: element.querySelector('.paragraph-marker')?.textContent })),
    textFrames: [...(currentHost?.querySelectorAll('.text-frame') || [])].map(element => ({ name: element.closest('.slide-element')?.title, clientHeight: element.clientHeight, scrollHeight: element.scrollHeight, whiteSpace: getComputedStyle(element).whiteSpace, writingMode: getComputedStyle(element).writingMode, textOrientation: getComputedStyle(element).textOrientation, direction: getComputedStyle(element).direction, paragraphDirection: getComputedStyle(element.querySelector('.text-paragraph')).direction, columnCount: getComputedStyle(element).columnCount, columnGap: getComputedStyle(element).columnGap, frameWidth: element.closest('.slide-frame')?.getBoundingClientRect().width, runRects: element.closest('.slide-element')?.title?.startsWith('Two-column') ? (() => { const range = document.createRange(); range.selectNodeContents(element.querySelector('.text-run')); return [...range.getClientRects()].map(rect => ({ left: rect.left, right: rect.right, top: rect.top })) })() : undefined, paragraphRects: element.closest('.slide-element')?.title === 'WordArt vertical RTL' ? [...element.querySelectorAll('.text-paragraph')].map(paragraph => { const rect = paragraph.getBoundingClientRect(); return { left: rect.left, top: rect.top } }) : undefined, borderStyle: getComputedStyle(element.closest('.slide-element')).borderStyle, borderWidth: getComputedStyle(element.closest('.slide-element')).borderWidth })),
    textRuns: [...(currentHost?.querySelectorAll('.text-run') || [])].map(element => ({ name: element.closest('.slide-element')?.title, text: element.textContent, textShadow: getComputedStyle(element).textShadow, textStrokeColor: getComputedStyle(element).webkitTextStrokeColor, textStrokeWidth: getComputedStyle(element).webkitTextStrokeWidth, visibility: getComputedStyle(element).visibility, opacity: getComputedStyle(element).opacity, color: getComputedStyle(element).color, fontSize: getComputedStyle(element).fontSize, fontFamily: getComputedStyle(element).fontFamily, fontWeight: getComputedStyle(element).fontWeight, fontStyle: getComputedStyle(element).fontStyle, textDecorationLine: getComputedStyle(element).textDecorationLine, verticalAlign: getComputedStyle(element).verticalAlign, maskImage: getComputedStyle(element).maskImage, wipeProgress: getComputedStyle(element).getPropertyValue('--wipe-progress').trim(), start: Number(element.dataset.charStart), end: Number(element.dataset.charEnd) })),
    tables: [...(currentHost?.querySelectorAll('.slide-table') || [])].map(table => ({
      cells: [...table.querySelectorAll('td')].map(cell => ({
        text: cell.textContent,
        colSpan: cell.colSpan,
        rowSpan: cell.rowSpan,
        backgroundColor: getComputedStyle(cell).backgroundColor,
        borderTopWidth: getComputedStyle(cell).borderTopWidth,
        autoFit: cell.dataset.autofit,
        clientHeight: cell.clientHeight,
        scrollHeight: cell.scrollHeight,
        fontSize: parseFloat(getComputedStyle(cell.querySelector('.text-run') || cell).fontSize),
        columnCount: getComputedStyle(cell.querySelector('.table-cell-text')).columnCount,
        columnGap: getComputedStyle(cell.querySelector('.table-cell-text')).columnGap,
        direction: getComputedStyle(cell.querySelector('.table-cell-text')).direction,
        paragraphDirection: getComputedStyle(cell.querySelector('.text-paragraph')).direction,
        frameWidth: cell.closest('.slide-frame')?.getBoundingClientRect().width,
        runRects: cell.textContent.includes('fills its first text column') ? (() => { const range = document.createRange(); range.selectNodeContents(cell.querySelector('.text-run')); return [...range.getClientRects()].map(rect => ({ left: rect.left, right: rect.right, top: rect.top })) })() : undefined,
      })),
    })),
    menuItems: [...document.querySelectorAll('.presentation-menu button')].map(item => item.textContent),
    screenOverlay: document.querySelector('.screen-overlay')?.className,
    stageClasses: [...(document.querySelector('.stage-shell')?.classList || [])],
    cutBlackObserved: window.__decklineCutBlackObserved || false,
    zoomTransform: getComputedStyle(currentHost).transform,
    slideTransition: (() => {
      const animation = currentHost?.getAnimations()[0]
      return animation ? {
        scale: new DOMMatrix(getComputedStyle(currentHost).transform).a,
        translateX: new DOMMatrix(getComputedStyle(currentHost).transform).e,
        translateY: new DOMMatrix(getComputedStyle(currentHost).transform).f,
        keyframes: animation.effect.getKeyframes().map(frame => frame.transform),
        opacityKeyframes: animation.effect.getKeyframes().map(frame => frame.opacity),
        clipKeyframes: animation.effect.getKeyframes().map(frame => frame.clipPath),
        clipPath: getComputedStyle(currentHost).clipPath,
        duration: animation.effect.getTiming().duration,
        currentTime: animation.currentTime,
      } : undefined
    })(),
    flashTransition: (() => {
      const overlay = currentHost?.querySelector('.flash-transition-overlay')
      const animation = overlay?.getAnimations()[0]
      return animation ? {
        opacityKeyframes: animation.effect.getKeyframes().map(frame => frame.opacity),
        duration: animation.effect.getTiming().duration,
        opacity: getComputedStyle(overlay).opacity,
        backgroundColor: getComputedStyle(overlay).backgroundColor,
      } : undefined
    })(),
    transitionUnderlay: { slideCount: outgoingHost ? 1 : 0, frame: outgoingRect && { left: outgoingRect.left, top: outgoingRect.top, width: outgoingRect.width, height: outgoingRect.height }, translateX: outgoingHost && new DOMMatrix(getComputedStyle(outgoingHost).transform).e, translateY: outgoingHost && new DOMMatrix(getComputedStyle(outgoingHost).transform).f, shapeNames: [...(outgoingHost?.querySelectorAll('.slide-element') || [])].map(element => element.title), pulling: underlay?.classList.contains('is-pulling'), splittingIn: underlay?.classList.contains('is-splitting-in'), layerZIndex: getComputedStyle(underlay).zIndex, hostZIndex: getComputedStyle(currentHost).zIndex, keyframes: outgoingHost?.getAnimations()[0]?.effect?.getKeyframes().map(frame => frame.transform), opacityKeyframes: outgoingHost?.getAnimations()[0]?.effect?.getKeyframes().map(frame => frame.opacity), clipKeyframes: outgoingHost?.getAnimations()[0]?.effect?.getKeyframes().map(frame => frame.clipPath), duration: outgoingHost?.getAnimations()[0]?.effect?.getTiming().duration },
    ink: {
      lines: [...(currentHost?.querySelectorAll('.ink-layer polyline') || [])].map(line => ({ stroke: line.getAttribute('stroke'), opacity: line.getAttribute('opacity') })),
      laserCount: currentHost?.querySelectorAll('.ink-layer circle').length || 0,
    },
    warnings: [...document.querySelectorAll('.warning-pill')].map(element => element.textContent),
    error: document.querySelector('.error-card')?.textContent?.replace(/\\s+/g, ' ').trim(),
  })})()`)
  const stateAtElapsed = async (start, elapsed) => {
    let snapshot = await state()
    while (snapshot.now - start < elapsed) {
      await delay(Math.min(elapsed - (snapshot.now - start), 25))
      snapshot = await state()
    }
    return snapshot
  }
  const shape = (snapshot, name) => snapshot.shapes.find(item => item.name === name)
  const rgbChannels = color => color.match(/\d+(?:\.\d+)?/g)?.slice(0, 3).map(Number) || []
  const transformValues = item => {
    const values = /matrix\(([^)]+)\)/.exec(item.transform)?.[1].split(',').map(Number)
    return values || [1, 0, 0, 1, 0, 0]
  }
  const rotation = item => Math.atan2(transformValues(item)[1], transformValues(item)[0]) * 180 / Math.PI
  const transformScale = item => Math.hypot(transformValues(item)[0], transformValues(item)[1])
  const paragraph = (snapshot, text) => snapshot.paragraphs.find(item => item.text === text)

  await call('Runtime.enable')
  await call('DOM.enable')
  await call('Page.enable')
  await call('Page.addScriptToEvaluateOnNewDocument', { source: `(() => {
    const NativeWorker = globalThis.Worker
    globalThis.Worker = class extends NativeWorker {
      constructor(...args) {
        super(...args)
        globalThis.__decklineWorkerCreated = true
      }
    }
    const nativeOpen = globalThis.open.bind(globalThis)
    globalThis.__decklineOpenedWindows = []
    globalThis.open = (...args) => {
      const opened = nativeOpen(...args)
      if (opened) globalThis.__decklineOpenedWindows.push(opened)
      return opened
    }
  })()` })
  await call('Page.reload', { ignoreCache: true })
  let inputObjectId
  for (let attempt = 0; attempt < 50; attempt++) {
    const input = await call('Runtime.evaluate', { expression: `document.querySelector('input[type=file]')` })
    inputObjectId = input.result.objectId
    if (inputObjectId) break
    await delay(100)
  }
  assert(inputObjectId, 'PPTX file input was not found.')
  let upload
  for (let attempt = 0; attempt < 10; attempt++) {
    const currentInput = await call('Runtime.evaluate', { expression: `document.querySelector('input[type=file]')` })
    inputObjectId = currentInput.result.objectId
    await call('DOM.setFileInputFiles', { objectId: inputObjectId, files: [inputPptx || fixturePath] })
    upload = await evaluate(`(() => { const input = document.querySelector('input[type=file]'); const result = { count: input.files?.length, name: input.files?.[0]?.name }; input.dispatchEvent(new Event('change', { bubbles: true })); return result })()`)
    await delay(200)
    if (await evaluate(`document.querySelector('.deck-title')?.textContent !== 'No presentation open' || Boolean(document.querySelector('.error-card'))`)) break
  }

  let initial
  for (let attempt = 0; attempt < 50; attempt++) {
    await delay(100)
    if (!await evaluate(`Boolean(document.querySelector('.stage-shell > .slide-host'))`)) {
      const error = await evaluate(`document.querySelector('.error-card')?.textContent?.replace(/\\s+/g, ' ').trim()`)
      if (error) { initial = { error, shapes: [] }; break }
      continue
    }
    initial = await state()
    if (captureOnly ? initial.slide && !initial.error : initial.shapes.length === 66) break
  }
  assert(await evaluate('globalThis.__decklineWorkerCreated === true'), 'PPTX parsing did not create a Web Worker.')
  assert(/^url\(["']?blob:/.test(initial.slideBackground?.image || ''), `An embedded slide background image must resolve to a local blob URL: ${JSON.stringify({ image: initial.slideBackground?.image, error: initial.error, warnings: initial.warnings, upload })}`)
  assert(initial.slideBackground?.size.startsWith('125% 100%'), `A cropped slide background image must scale to its source rectangle: ${initial.slideBackground?.size}`)
  const backgroundLoaded = await evaluate(`new Promise(resolve => {
    const frame = document.querySelector('.stage-shell > .slide-host .slide-frame')
    const source = /^url\\(["']?(.*?)["']?\\)$/.exec(frame?.style.backgroundImage || '')?.[1]
    if (!source) return resolve(false)
    const image = new Image()
    image.onload = () => resolve(image.naturalWidth === 1)
    image.onerror = () => resolve(false)
    image.src = source
  })`)
  assert(backgroundLoaded, 'The embedded slide background image must decode in the browser.')
  if (captureOnly) {
    assert(initial.slide && !initial.error, `The PPTX did not render: ${JSON.stringify({ upload, initial, runtimeErrors })}`)
    const textLayout = await evaluate(`(() => [...document.querySelectorAll('.stage-shell > .slide-host .slide-element')].map(shape => {
      const text = shape.querySelector('.text-frame')
      const run = text?.querySelector('.text-run')
      if (!text || !run) return null
      const rect = node => { const { top, bottom, height } = node.getBoundingClientRect(); return { top, bottom, height } }
      return { name: shape.title, shape: rect(shape), textFrame: { ...rect(text), clientHeight: text.clientHeight, scrollHeight: text.scrollHeight, overflow: getComputedStyle(text).overflow }, paragraph: rect(run.parentElement), run: rect(run), fontSize: getComputedStyle(run).fontSize, bodyFit: shape.className }
    }).filter(Boolean))()`)
    const rect = await evaluate(`(() => { const { x, y, width, height } = document.querySelector('.stage-shell > .slide-host .slide-frame').getBoundingClientRect(); return { x, y, width, height } })()`)
    const screenshot = await call('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true, clip: { ...rect, scale: 1280 / rect.width } })
    assert(process.env.DECKLINE_CAPTURE_SLIDE_PATH, 'Set DECKLINE_CAPTURE_SLIDE_PATH for the browser capture.')
    await writeFile(process.env.DECKLINE_CAPTURE_SLIDE_PATH, Buffer.from(screenshot.data, 'base64'))
    console.log(JSON.stringify({ screenshot: process.env.DECKLINE_CAPTURE_SLIDE_PATH, slide: initial.slide, hostFrame: initial.hostFrame, shapes: initial.shapes, paragraphs: initial.paragraphs, textRuns: initial.textRuns, textLayout }))
  } else {
  assert(initial.shapes.length === 80, `The animation fixture did not render: ${JSON.stringify({ upload, initial, runtimeErrors })}`)
  const textOverflow = await evaluate(`(() => {
    const clip = document.querySelector('.stage-shell > .slide-host [title="Picture bullet"] .text-frame')
    const ellipsis = document.querySelector('.stage-shell > .slide-host [title="Autofit reference"] .text-frame')
    const clippedStyle = getComputedStyle(clip)
    const ellipsisStyle = getComputedStyle(ellipsis)
    const run = clip.querySelector('.text-run')
    return { horizontal: clippedStyle.overflowX, vertical: clippedStyle.overflowY, textLength: run.textContent.length, whiteSpace: clippedStyle.whiteSpace, ellipsisVertical: ellipsisStyle.overflowY }
  })()`)
  assert(textOverflow.horizontal === 'clip' && textOverflow.vertical === 'clip'
    && textOverflow.textLength > 100 && textOverflow.whiteSpace === 'nowrap' && ['clip', 'hidden'].includes(textOverflow.ellipsisVertical)
    && initial.warnings.some(warning => warning.includes('vertOverflow="ellipsis"')),
    `DrawingML text overflow must clip at the text frame and warn when ellipsis is approximated: ${JSON.stringify({ textOverflow, warnings: initial.warnings })}`)
  const initialThumbnailWindow = await evaluate(`(() => {
    const list = document.querySelector('.thumb-list')
    const items = [...(list?.querySelectorAll('.thumb-button') || [])]
    return { total: Number(document.querySelector('.slide-count')?.textContent), rendered: items.length, first: Number(items[0]?.dataset.thumbnailIndex), last: Number(items.at(-1)?.dataset.thumbnailIndex) }
  })()`)
  assert(initialThumbnailWindow.total === 67 && initialThumbnailWindow.rendered > 0 && initialThumbnailWindow.rendered < initialThumbnailWindow.total,
    `The slide sidebar must render a bounded thumbnail window: ${JSON.stringify(initialThumbnailWindow)}`)
  const autofitMetrics = await evaluate(`(() => {
    const fit = document.querySelector('.stage-shell > .slide-host .slide-element[title="Normal autofit"]')
    const reference = document.querySelector('.stage-shell > .slide-host .slide-element[title="Autofit reference"]')
    const run = fit?.querySelector('.text-run')
    const paragraph = run?.parentElement
    const css = element => getComputedStyle(element)
    return { fontScale: Number.parseFloat(css(run).fontSize) / Number.parseFloat(css(reference?.querySelector('.text-run')).fontSize), lineHeightRatio: Number.parseFloat(css(paragraph).lineHeight) / Number.parseFloat(css(paragraph).fontSize) }
  })()`)
  assert(Math.abs(autofitMetrics.fontScale - 0.8) < 0.01 && Math.abs(autofitMetrics.lineHeightRatio - 1.5) < 0.01,
    `a:normAutofit must scale each text run and reduce percentage line spacing: ${JSON.stringify(autofitMetrics)}`)
  const dynamicAutoFitMetrics = await evaluate(`(() => {
    const shape = document.querySelector('.stage-shell > .slide-host .slide-element[title="Normal autofit dynamic"]')
    const baseline = document.querySelector('.stage-shell > .slide-host .slide-element[title="Shape autofit baseline"]')
    const frame = shape?.querySelector('.text-frame')
    const run = frame?.querySelector('.text-run')
    const baselineRun = baseline?.querySelector('.text-run')
    return { scale: Number.parseFloat(getComputedStyle(run).fontSize) / Number.parseFloat(getComputedStyle(baselineRun).fontSize), scrollWidth: frame?.scrollWidth, clientWidth: frame?.clientWidth, scrollHeight: frame?.scrollHeight, clientHeight: frame?.clientHeight }
  })()`)
  assert(dynamicAutoFitMetrics.scale < 0.99 && dynamicAutoFitMetrics.scrollWidth <= dynamicAutoFitMetrics.clientWidth && dynamicAutoFitMetrics.scrollHeight <= dynamicAutoFitMetrics.clientHeight,
    `a:normAutofit without a stored scale must shrink overflow text to fit its frame: ${JSON.stringify(dynamicAutoFitMetrics)}`)
  const shapeAutoFitMetrics = await evaluate(`(() => {
    const fit = document.querySelector('.stage-shell > .slide-host .slide-element[title="Shape autofit"]')
    const baseline = document.querySelector('.stage-shell > .slide-host .slide-element[title="Shape autofit baseline"]')
    const frame = fit?.querySelector('.text-frame')
    return { fitHeight: fit?.getBoundingClientRect().height, baselineHeight: baseline?.getBoundingClientRect().height, scrollHeight: frame?.scrollHeight, clientHeight: frame?.clientHeight }
  })()`)
  assert(shapeAutoFitMetrics.fitHeight > shapeAutoFitMetrics.baselineHeight + 8
    && shapeAutoFitMetrics.clientHeight >= shapeAutoFitMetrics.scrollHeight,
    `a:spAutoFit must expand the shape enough to contain overflow: ${JSON.stringify(shapeAutoFitMetrics)}`)
  const renderedLine = shape(initial, 'Stroke color').line
  assert(renderedLine?.x1 === 0 && renderedLine.y1 === 0 && renderedLine.x2 === 1000 && renderedLine.y2 === 1000
    && renderedLine.markerStart === 'url(#pptx-line-42-head)' && renderedLine.markerEnd === 'url(#pptx-line-42-tail)'
    && renderedLine.headPath === 'M0 0L10 5L0 10Z' && renderedLine.tailPath === 'M0 5L5 0L10 5L5 10Z',
    `A DrawingML line must preserve its diagonal endpoints and sized head/tail markers: ${JSON.stringify(renderedLine)}`)
  assert(shape(initial, 'Slide entrance and exit').visibility === 'hidden', 'A slide-filter entrance must start off-stage and hidden before its trigger.')
  const customGeometry = shape(initial, 'Custom geometry').customPaths[0]
  assert(customGeometry?.d.startsWith('M') && customGeometry.d.includes('C') && customGeometry.fill === 'rgb(163, 230, 53)' && customGeometry.stroke === 'rgb(22, 101, 52)' && Number.parseFloat(customGeometry.strokeWidth) > 0, `Custom geometry cubic paths must preserve their path, fill, and outline: ${JSON.stringify(customGeometry)}`)
  const customGeometryDash = await evaluate(`(() => { const path = document.querySelector('.stage-shell > .slide-host .slide-element[title="Custom geometry"] .custom-geometry path'); const style = path && getComputedStyle(path); const width = Number.parseFloat(style?.strokeWidth || '0'); return { dashRatios: (style?.strokeDasharray || '').split(/[ ,]+/).filter(Boolean).map(value => Number.parseFloat(value) / width), cap: style?.strokeLinecap } })()`)
  assert(customGeometryDash.dashRatios.length === 4 && [2.5, 0.5, 0.5, 0.25].every((ratio, index) => Math.abs(customGeometryDash.dashRatios[index] - ratio) < 0.02)
    && customGeometryDash.cap === 'round'
    && !initial.warnings.some(warning => warning.includes('Custom dash ratios on preset shape outlines')),
    `DrawingML custom dash ratios and rounded caps must render on custom SVG geometry: ${JSON.stringify({ ...customGeometryDash, warnings: initial.warnings })}`)
  const customQuadratic = shape(initial, 'Custom geometry').customPaths[1]
  assert(customQuadratic?.d.includes('Q') && customQuadratic.fill === 'none' && customQuadratic.stroke === 'rgb(22, 101, 52)', `Custom geometry must preserve a second quadratic path and its no-fill mode: ${JSON.stringify(customQuadratic)}`)
  const customArc = shape(initial, 'Custom geometry').customPaths[2]
  assert(customArc?.d.includes('A 400 250 0 0 1 900') && customArc.fill === 'none' && customArc.stroke === 'rgb(22, 101, 52)', `DrawingML arcTo must render as a clockwise SVG ellipse arc: ${JSON.stringify(customArc)}`)
  const guideDrivenPath = shape(initial, 'Custom geometry').customPaths[3]
  assert(guideDrivenPath?.d.includes('M 250 750 L 750 750 L 750 250 Z'), `Custom geometry points must resolve chained guides, adjustment values, and standard formulas: ${JSON.stringify(guideDrivenPath)}`)
  const guideFormulaPath = shape(initial, 'Custom geometry').customPaths[4]
  assert(guideFormulaPath?.d.includes('M 200 50 L 10 20 10 13 100 100 25 0 0 16 0 0 1000 250'), `Custom geometry points must resolve the supported guide formula operators: ${JSON.stringify(guideFormulaPath)}`)
  const guideFormulaArcPath = shape(initial, 'Custom geometry').customPaths[5]
  assert(guideFormulaArcPath?.d.includes('A 200 200 0 0 1'), `Custom geometry arc angles must resolve guide formulas and built-in angle constants: ${JSON.stringify(guideFormulaArcPath)}`)
  const customTextRectMetrics = await evaluate(`(() => {
    const read = name => {
      const shape = document.querySelector('.stage-shell > .slide-host .slide-element[title="' + name + '"]')
      const frame = shape?.querySelector('.text-frame')
      const outer = shape?.getBoundingClientRect(), text = frame?.getBoundingClientRect()
      return outer && text ? { left: (text.left - outer.left) / outer.width, top: (text.top - outer.top) / outer.height, right: (text.right - outer.left) / outer.width, bottom: (text.bottom - outer.top) / outer.height } : undefined
    }
    return { custom: read('Custom text rectangle'), invalid: read('Invalid custom text rectangle') }
  })()`)
  assert(customTextRectMetrics.custom && Math.abs(customTextRectMetrics.custom.left - 0.25) < 0.01
    && Math.abs(customTextRectMetrics.custom.top - 0.2) < 0.01 && Math.abs(customTextRectMetrics.custom.right - 0.9) < 0.01
    && Math.abs(customTextRectMetrics.custom.bottom - 0.8) < 0.01
    && customTextRectMetrics.invalid && Math.abs(customTextRectMetrics.invalid.left) < 0.01 && Math.abs(customTextRectMetrics.invalid.top) < 0.01
    && Math.abs(customTextRectMetrics.invalid.right - 1) < 0.01 && Math.abs(customTextRectMetrics.invalid.bottom - 1) < 0.01
    && initial.warnings.some(warning => warning.includes('custom geometry text rectangle has invalid bounds')),
  `Custom geometry text rectangles must resolve guide-based bounds and fall back safely when invalid: ${JSON.stringify({ customTextRectMetrics, warnings: initial.warnings })}`)
  const customGradientMetrics = await evaluate(`(() => {
    const path = document.querySelector('.stage-shell > .slide-host .slide-element[title="Custom gradient geometry"] .custom-geometry path')
    const read = property => {
      const paint = path?.getAttribute(property) || ''
      const id = /^url\\(#(.+)\\)$/.exec(paint)?.[1]
      const gradient = id ? document.getElementById(id) : undefined
      return { paint, x1: Number(gradient?.getAttribute('x1')), y1: Number(gradient?.getAttribute('y1')), x2: Number(gradient?.getAttribute('x2')), y2: Number(gradient?.getAttribute('y2')), units: gradient?.getAttribute('gradientUnits'), stops: [...(gradient?.querySelectorAll('stop') || [])].map(stop => ({ offset: stop.getAttribute('offset'), color: stop.getAttribute('stop-color'), opacity: stop.getAttribute('stop-opacity') })) }
    }
    return { fill: read('fill'), stroke: read('stroke') }
  })()`)
  assert(customGradientMetrics.fill.units === 'userSpaceOnUse' && customGradientMetrics.fill.stops.length === 3
    && customGradientMetrics.fill.stops[0]?.offset === '0%' && customGradientMetrics.fill.stops[1]?.offset === '50%' && customGradientMetrics.fill.stops[2]?.offset === '100%'
    && customGradientMetrics.fill.stops[0]?.color === '#ff0000' && customGradientMetrics.fill.stops[1]?.color === 'rgb(0, 255, 0)' && customGradientMetrics.fill.stops[1]?.opacity === '0.5' && customGradientMetrics.fill.stops[2]?.color === '#0000ff'
    && Math.abs(customGradientMetrics.fill.x1) < 0.01 && Math.abs(customGradientMetrics.fill.y1 - 500) < 0.01 && Math.abs(customGradientMetrics.fill.x2 - 1000) < 0.01 && Math.abs(customGradientMetrics.fill.y2 - 500) < 0.01
    && customGradientMetrics.stroke.stops.length === 2 && Math.abs(customGradientMetrics.stroke.x1 - 500) < 0.01 && Math.abs(customGradientMetrics.stroke.y1) < 0.01 && Math.abs(customGradientMetrics.stroke.x2 - 500) < 0.01 && Math.abs(customGradientMetrics.stroke.y2 - 1000) < 0.01,
  `Custom geometry linear fills and outlines must retain DrawingML gradient directions, stops, and alpha: ${JSON.stringify(customGradientMetrics)}`)
  assert(!initial.warnings.some(warning => warning.includes('custom geometry path segments')), `Valid custom-geometry arcs must not be omitted: ${JSON.stringify(initial.warnings)}`)
  const motionStart = shape(initial, 'Motion path').frameRect.left / shape(initial, 'Motion path').frameRect.frameWidth
  const relativeMotionStart = shape(initial, 'Relative position motion').frameRect.left / shape(initial, 'Relative position motion').frameRect.frameWidth
  const fromToMotionStart = shape(initial, 'From/to motion').frameRect.left / shape(initial, 'From/to motion').frameRect.frameWidth
  const fromByMotionStart = shape(initial, 'From/by motion').frameRect.left / shape(initial, 'From/by motion').frameRect.frameWidth
  const toMotionStart = shape(initial, 'To-only motion').frameRect.left / shape(initial, 'To-only motion').frameRect.frameWidth
  const fixedMotionStart = shape(initial, 'Fixed motion path').frameRect.left / shape(initial, 'Fixed motion path').frameRect.frameWidth
  const rotatedMotionStart = shape(initial, 'Rotated fixed motion path').frameRect
  const curveStart = shape(initial, 'Curved motion path').frameRect
  const formulaMotionStart = shape(initial, 'Formula fixed motion path').frameRect
  const parentMotionStart = shape(initial, 'Parent origin grouped motion path').frameRect
  const nestedParentMotionStart = shape(initial, 'Nested group child').frameRect
  assert(!initial.warnings.some(warning => warning.includes('animation-progress ($) formulas')), `Supported animation-progress motion formulas must not be reported as unsupported: ${JSON.stringify(initial.warnings)}`)
  assert(!initial.warnings.some(warning => warning.includes('coordinate formula was skipped')), `Supported static motion-path formulas must parse: ${JSON.stringify(initial.warnings)}`)
  assert(initial.textRuns.find(run => run.text === 'Built paragraph').textShadow !== 'none', 'Text outer shadows must render.')
  const staticStrikethroughRuns = initial.textRuns.filter(run => run.name === 'Text range opacity' && run.text.trim())
  const staticStrikeStyles = await evaluate(`(() => [...(document.querySelector('.stage-shell > .slide-host')?.querySelectorAll('.slide-element[title="Text range opacity"] .text-run') || [])].filter(run => run.textContent.trim()).map(run => getComputedStyle(run).textDecorationStyle))()`)
  assert(staticStrikethroughRuns.length > 0 && staticStrikethroughRuns.every(run => run.textDecorationLine.includes('line-through')) && staticStrikeStyles.length === staticStrikethroughRuns.length && staticStrikeStyles.every(style => style === 'double'), `DrawingML a:rPr strike="dblStrike" must render as a double CSS line-through across text-run segmentation: ${JSON.stringify({ runs: staticStrikethroughRuns, styles: staticStrikeStyles })}`)
  const textUnderlineStyles = await evaluate(`(() => Object.fromEntries([...(document.querySelector('.stage-shell > .slide-host')?.querySelectorAll('.slide-element[title="Opacity formula targets"] .text-run') || [])].filter(run => run.textContent.trim()).map(run => [run.textContent.trim(), getComputedStyle(run).textDecorationStyle])))()`)
  assert(textUnderlineStyles.Run === 'dotted' && textUnderlineStyles.Dash === 'dashed' && textUnderlineStyles.Wave === 'wavy' && textUnderlineStyles.Pair === 'double' && textUnderlineStyles.Word === 'solid' && textUnderlineStyles.Mix === 'dotted', `DrawingML underline styles must map to CSS patterns and keep underline precedence when one element combines different underline/strike styles: ${JSON.stringify(textUnderlineStyles)}`)
  assert(initial.warnings.some(warning => warning.includes('underline value "words" is approximated')) && initial.warnings.some(warning => warning.includes('different underline and strikethrough styles')), `Unrepresentable underline variants and conflicting CSS decoration styles must be reported: ${JSON.stringify(initial.warnings)}`)
  const outlinedParagraphRuns = await evaluate(`(() => [...(document.querySelector('.stage-shell > .slide-host')?.querySelectorAll('.slide-element[title="Paragraph wipe target"] .text-run') || [])].map(run => ({ color: getComputedStyle(run).webkitTextStrokeColor, width: parseFloat(getComputedStyle(run).webkitTextStrokeWidth) })))()`)
  assert(outlinedParagraphRuns.length === 2 && outlinedParagraphRuns.every(run => run.color === 'rgb(255, 0, 0)' && run.width > 0), `A solid DrawingML text outline must render with its color and width across p:charRg segmentation: ${JSON.stringify(outlinedParagraphRuns)}`)
  assert(['Al', '\n', 'ways'].every(text => initial.textRuns.some(run => run.text === text && run.visibility === 'hidden')) && initial.textRuns.find(run => run.text === 'Al')?.start === 16 && initial.textRuns.find(run => run.text === 'ways')?.end === 23 && initial.textRuns.some(run => run.text === ' visible' && run.visibility === 'visible'), 'Character ranges must isolate the half-open selected text interval across styled spans.')
  assert(shape(initial, 'Automatic fade').visibility === 'hidden' && shape(initial, 'Automatic fade').opacity === '0', `An automatic entrance must remain hidden throughout its delay: ${JSON.stringify({ shape: shape(initial, 'Automatic fade'), warnings: initial.warnings, reducedMotion: await evaluate("matchMedia('(prefers-reduced-motion: reduce)').matches") })}`)
  const motionWarnings = initial.warnings.filter(warning => warning.toLowerCase().includes('motion'))
  assert(!motionWarnings.some(warning => /empty motion path/i.test(warning)), `A standards-compliant empty motion path must retain its timing without an empty-path warning: ${JSON.stringify(motionWarnings)}`)
  assert(initial.warnings.some(warning => warning.includes('style.fontSize p:set animation with an unsupported target was skipped')), `An invalid p:set font-size target must produce a warning: ${JSON.stringify(initial.warnings)}`)
  assert(initial.warnings.some(warning => warning.includes('A generic text style animation with unsupported values, formulas, or interpolation was skipped')), `An unsafe font-family animation must be rejected with a warning: ${JSON.stringify(initial.warnings)}`)
  const fontSizeSetRuns = await evaluate(`(() => [...(document.querySelector('.stage-shell > .slide-host')?.querySelectorAll('.slide-element[title="Paragraph wipe target"] .text-run') || [])].map(run => ({ start: Number(run.dataset.charStart), end: Number(run.dataset.charEnd), size: parseFloat(getComputedStyle(run).fontSize) })))()`)
  assert(fontSizeSetRuns.length === 2 && fontSizeSetRuns[0]?.start === 0 && fontSizeSetRuns[0]?.end === 15 && fontSizeSetRuns[1]?.start === 15 && fontSizeSetRuns[1]?.end === 21 && Math.abs(fontSizeSetRuns[1]?.size / fontSizeSetRuns[0]?.size - 1.5) < 0.05, `An automatic p:set style.fontSize must hold a fixed factor on only its p:charRg range: ${JSON.stringify(fontSizeSetRuns)}`)
  const setBoldWeight = await evaluate(`(() => [...document.querySelectorAll('.stage-shell > .slide-host .slide-element[title="By letter"] .text-run')].map(run => getComputedStyle(run).fontWeight))()`)
  assert(setBoldWeight.length > 0 && setBoldWeight.every(weight => weight === '700'), `A p:set style.fontWeight bold preset must apply to its text target: ${JSON.stringify(setBoldWeight)}`)
  const resetWeight = initial.textRuns.find(run => run.text === 'ways')?.fontWeight
  assert(resetWeight === '700', `A p:set style.fontWeight none preset must restore the targeted run's static bold weight: ${resetWeight}`)
  const pictureBullet = await evaluate(`(() => { const image = document.querySelector('.stage-shell > .slide-host .slide-element[title="Picture bullet"] .paragraph-picture-bullet'); const text = image?.parentElement?.querySelector('.text-run'); return image ? { src: image.getAttribute('src'), width: parseFloat(getComputedStyle(image).width), height: parseFloat(getComputedStyle(image).height), textSize: parseFloat(getComputedStyle(text).fontSize) } : null })()`)
  assert(pictureBullet?.src?.startsWith('blob:') && pictureBullet.width > 0 && pictureBullet.height > 0 && Math.abs(pictureBullet.width / pictureBullet.textSize - 1.25) < 0.05, `An a:buBlip picture bullet must resolve inline at its 125% buSzPct size: ${JSON.stringify(pictureBullet)}`)
  assert(!initial.warnings.some(warning => warning.includes('shape iteration')), `Shape iteration must be handled rather than reported as unsupported: ${JSON.stringify(initial.warnings)}`)
  assert(shape(initial, 'Click fade').visibility === 'hidden' && shape(initial, 'With previous').visibility === 'hidden', 'Entrance targets must start hidden.')
  assert(shape(initial, 'Exit after previous').visibility === 'visible' && shape(initial, 'Second click').visibility === 'hidden', 'Exit targets must start visible while later entrances remain hidden.')
  assert(shape(initial, 'Wipe exit').visibility === 'visible', 'A wipe exit target must begin visible.')
  assert(shape(initial, 'Blinds animation').maskImage.includes('repeating-linear-gradient') && shape(initial, 'Blinds animation').blindsStripeSize === '0%', 'An entrance blinds animation must begin with all slats closed.')
  const checkerStart = shape(initial, 'Checker animation').clipPath
  assert(checkerStart.startsWith('path('), `A checkerboard entrance must begin as a tiled object mask: ${checkerStart}`)
  assert(['W', 'O', 'R', 'D'].every(text => initial.textRuns.find(run => run.text === text)?.visibility === 'hidden'), 'By-letter animation text must start hidden before its click step.')
  assert(['ONE', 'TWO'].every(text => initial.textRuns.find(run => run.text === text)?.visibility === 'hidden'), 'By-word animation text must start hidden before its click step.')
  assert(paragraph(initial, 'Built paragraph').visibility === 'hidden' && paragraph(initial, 'Al\nways visible').visibility === 'visible', 'Only the targeted paragraph range must start hidden.')

  assert(initial.textRuns.find(run => run.name === 'Character slide animation' && run.text === 'REST')?.visibility === 'visible'
    && initial.textRuns.find(run => run.name === 'Character slide animation' && run.text === 'MOVE')?.visibility === 'hidden'
    && !initial.warnings.some(warning => warning.includes('Text-range slide')),
  'A valid character-range slide entrance must hide only its selected range without an unsupported-effect warning.')
  const autoActive = await evaluate(`new Promise(resolve => {
    const started = performance.now()
    const checkFrame = () => {
      const element = document.querySelector('.stage-shell > .slide-host [title="Automatic fade"]')
      const opacity = Number(element && getComputedStyle(element).opacity)
      if (element && getComputedStyle(element).visibility === 'visible' && opacity > 0 && opacity < 0.3) resolve(true)
      else if (performance.now() - started > 12000) resolve(false)
      else requestAnimationFrame(checkFrame)
    }
    checkFrame()
  })`)
  assert(autoActive, `The repeated automatic fade must enter its first interpolation after its delay: ${JSON.stringify({ automatic: shape(await state(), 'Automatic fade'), warnings: initial.warnings })}`)
  await delay(200)
  const autoRepeat = await state()
  assert(Number(shape(autoRepeat, 'Automatic fade').opacity) > 0 && Number(shape(autoRepeat, 'Automatic fade').opacity) < 1, 'A repeated fade must restart its interpolation for the second iteration.')
  await delay(200)
  const autoComplete = await state()
  assert(shape(autoComplete, 'Automatic fade').visibility === 'visible' && shape(autoComplete, 'Automatic fade').opacity === '1', 'The repeated automatic sequence must hold its final state without a click.')

  const clickStartedAt = await evaluate(`(() => { const started = performance.now(); document.querySelector('.stage-shell').click(); return started })()`)
  const afterStart = await stateAtElapsed(clickStartedAt, 50)
  assert(shape(afterStart, 'Click fade').visibility === 'visible' && shape(afterStart, 'With previous').visibility === 'hidden', `The delayed with-effect must not start with the click effect: ${JSON.stringify({ click: shape(afterStart, 'Click fade'), delayed: shape(afterStart, 'With previous') })}`)
  assert(Number(shape(afterStart, 'Click fade').opacity) > 0 && Number(shape(afterStart, 'Click fade').opacity) < 0.29, `A fade with equal acceleration and deceleration should still be in the early part of its normalized time curve: ${shape(afterStart, 'Click fade').opacity}`)
  assert(Number(shape(afterStart, 'Parent eased fade').opacity) > 0 && Number(shape(afterStart, 'Parent eased fade').opacity) < 0.29, `A finite parent time-node curve must shape its child animation progress: ${shape(afterStart, 'Parent eased fade').opacity}`)
  assert(shape(afterStart, 'Parent timing after-effect').visibility === 'hidden', 'An after-effect must wait for the finite parent time-container duration.')
  const characterSlideDuring = await evaluate(`(() => { const shape = document.querySelector('.stage-shell > .slide-host .slide-element[title="Character slide animation"]'); const read = text => { const run = [...shape.querySelectorAll('.text-run')].find(item => item.textContent === text); const style = getComputedStyle(run); const matrix = style.transform === 'none' ? new DOMMatrix() : new DOMMatrix(style.transform); return { translateX: matrix.e, visibility: style.visibility } }; return { rest: read('REST'), move: read('MOVE'), frameWidth: shape.closest('.slide-frame')?.getBoundingClientRect().width } })()`)
  assert(characterSlideDuring.rest.visibility === 'visible' && Math.abs(characterSlideDuring.rest.translateX) < 0.01
    && characterSlideDuring.move.visibility === 'visible' && characterSlideDuring.move.translateX < 0 && characterSlideDuring.move.translateX > -characterSlideDuring.frameWidth,
    `A slide(fromLeft) character range must move only its selected text in from the stage edge: ${JSON.stringify(characterSlideDuring)}`)
  assert(Number.parseFloat(shape(afterStart, 'Blinds animation').blindsStripeSize) > 0 && Number.parseFloat(shape(afterStart, 'Blinds animation').blindsStripeSize) < 12.5, 'An object blinds entrance must open its horizontal mask strips over time.')
  assert(shape(afterStart, 'Checker animation').clipPath !== checkerStart, 'An object checkerboard must reveal tiles during its click animation.')
  const characterFades = afterStart.textRuns.filter(run => ['Al', '\n', 'ways'].includes(run.text))
  assert(characterFades.length === 3 && characterFades.every(run => run.visibility === 'visible' && Number(run.opacity) > 0 && Number(run.opacity) < 1), `The character-range fade must interpolate only the selected text across rich-text runs: ${JSON.stringify({ characterFades, warnings: afterStart.warnings })}`)
  const paragraphWipe = afterStart.paragraphs.find(item => item.text === 'Paragraph wipe target')
  assert(paragraphWipe?.maskImage.includes('linear-gradient') && Number.parseFloat(paragraphWipe.wipeProgress) > 0 && Number.parseFloat(paragraphWipe.wipeProgress) < 100 && paragraphWipe.opacity === '1' && shape(afterStart, 'Paragraph wipe target').maskImage === 'none', `A paragraph-range wipe must mask only the selected paragraph and preserve text opacity: ${JSON.stringify(paragraphWipe)}`)
  const letterRuns = afterStart.textRuns.filter(run => ['W', 'O', 'R', 'D'].includes(run.text))
  assert(letterRuns.length === 4 && Number(letterRuns[0].opacity) > 0 && Number(letterRuns[0].opacity) < 1 && letterRuns.slice(1).every(run => run.visibility === 'hidden'), `The tmAbs by-letter iteration must stagger grapheme starts: ${JSON.stringify(letterRuns)}`)
  const wordRuns = afterStart.textRuns.filter(run => run.name === 'By word' && ['ONE', 'TWO', ' '].includes(run.text))
  assert(wordRuns.length === 3 && Number(wordRuns.find(run => run.text === 'TWO')?.opacity) > 0 && wordRuns.find(run => run.text === 'ONE')?.visibility === 'hidden' && wordRuns.find(run => run.text === ' ')?.visibility === 'visible', `The backwards tmPct by-word iteration must reveal words in reverse and leave separators alone: ${JSON.stringify(wordRuns)}`)
  const baseFontSize = Number.parseFloat(initial.textRuns.find(run => run.text === 'ways')?.fontSize || '')
  const activeFontSize = Number.parseFloat(afterStart.textRuns.find(run => run.text === 'ways')?.fontSize || '')
  assert(activeFontSize > baseFontSize && activeFontSize < baseFontSize * 1.5, `A generic style.fontSize animation must interpolate on just the targeted character range: ${JSON.stringify({ baseFontSize, activeFontSize, targetRun: afterStart.textRuns.find(run => run.text === 'ways'), warnings: afterStart.warnings.filter(warning => /font.?size|animation/i.test(warning)) })}`)
  const fadeOpacity = Number(shape(afterStart, 'Click fade').opacity)
  assert(fadeOpacity > 0 && fadeOpacity < 1, 'A fade entrance must interpolate during its duration.')
  const motionAtStart = shape(afterStart, 'Motion path').frameRect.left / shape(afterStart, 'Motion path').frameRect.frameWidth
  assert(motionAtStart > motionStart + 0.2 && motionAtStart < motionStart + 0.25, `A later motion-path M must snap to its new subpath and continue interpolating there: ${motionAtStart - motionStart}`)
  const curveAtStart = shape(afterStart, 'Curved motion path').frameRect
  assert(curveAtStart.top / curveAtStart.frameHeight > curveStart.top / curveStart.frameHeight + 0.01, 'A cubic Bezier path must interpolate along its curved y-coordinate.')
  assert(shape(afterStart, 'Exit after previous').visibility === 'visible', 'The after-effect must wait for preceding effects.')
  for (const direction of ['right', 'left', 'up', 'down']) {
    const wipe = shape(afterStart, `Wipe ${direction}`)
    assert(wipe.visibility === 'hidden' && wipe.opacity === '1' && wipe.wipeProgress === '0%', `The ${direction} wipe entrance must start hidden behind a fully closed mask without fading.`)
  }

  const afterWith = await stateAtElapsed(clickStartedAt, 150)
  assert(Number(shape(afterWith, 'Parent speed delayed').opacity) > 0.6, `A positive parent time-node speed must scale its child's delay: ${shape(afterWith, 'Parent speed delayed').opacity}`)
  assert(shape(afterWith, 'With previous').visibility === 'visible' && shape(afterWith, 'Exit after previous').visibility === 'visible', 'The with-effect delay must be applied inside the click step.')

  const midIteration = await stateAtElapsed(clickStartedAt, 350)
  assert(midIteration.textRuns.find(run => run.text === 'D')?.visibility === 'hidden' && shape(midIteration, 'After letters').opacity === '1', `The delayed exit must not start before the final by-letter iteration: ${JSON.stringify({ lastLetter: midIteration.textRuns.find(run => run.text === 'D'), afterEffect: shape(midIteration, 'After letters') })}`)
  const beforeParentEnd = await stateAtElapsed(clickStartedAt, 500)
  assert(shape(beforeParentEnd, 'Parent timing after-effect').visibility === 'hidden', 'The parent time-container after-effect must remain pending before its duration ends.')
  const afterChain = await stateAtElapsed(clickStartedAt, 1000)
  const characterSlideComplete = await evaluate(`(() => { const shape = document.querySelector('.stage-shell > .slide-host .slide-element[title="Character slide animation"]'); return [...shape.querySelectorAll('.text-run')].map(run => { const style = getComputedStyle(run); const matrix = style.transform === 'none' ? new DOMMatrix() : new DOMMatrix(style.transform); return { text: run.textContent, translateX: matrix.e, visibility: style.visibility } }) })()`)
  assert(characterSlideComplete.length === 2 && characterSlideComplete.every(run => run.visibility === 'visible' && Math.abs(run.translateX) < 0.01),
    `A completed character-range slide entrance must reveal both ranges at their base positions: ${JSON.stringify(characterSlideComplete)}`)
  assert(shape(afterChain, 'Blinds animation').blindsStripeSize === '12.5%', 'The object blinds animation must finish with all mask strips open.')
  assert(shape(afterChain, 'Checker animation').clipPath !== checkerStart, 'The object checkerboard must hold its final fully open tile mask.')
  assert(shape(afterChain, 'After letters').visibility === 'hidden', 'The chained after-effect must start after the full by-letter duration.')
  assert(['Al', '\n', 'ways'].every(text => afterChain.textRuns.find(run => run.text === text)?.color === 'rgb(255, 0, 0)'), 'Character-range text color animation must preserve the selection and reach its final color.')
  assert(Math.abs(Number.parseFloat(afterChain.textRuns.find(run => run.text === 'ways')?.fontSize || '') - baseFontSize * 1.5) < 0.5, 'A held generic style.fontSize animation must retain its final scale.')
  assert(shape(afterChain, 'Click fade').visibility === 'visible' && shape(afterChain, 'With previous').visibility === 'visible', 'Both entrance effects must finish in the first click step.')
  assert(shape(afterChain, 'Parent timing after-effect').visibility === 'visible' && shape(afterChain, 'Parent timing after-effect').opacity === '1', 'The after-effect must begin and finish after the complete parent time-container duration.')
  assert(Math.abs(shape(afterChain, 'Motion path').frameRect.left / shape(afterChain, 'Motion path').frameRect.frameWidth - motionStart - 0.25) < 0.01, 'A straight layout-origin motion path must end at its OOXML slide-size offset.')
  const relativeMotionEnd = shape(afterChain, 'Relative position motion').frameRect.left / shape(afterChain, 'Relative position motion').frameRect.frameWidth
  assert(Math.abs(relativeMotionEnd - relativeMotionStart - 0.25) < 0.01, 'A parent-origin by-only motion must add its ST_Percentage offset to the object position.')
  assert(Math.abs(shape(afterChain, 'From/to motion').frameRect.left / shape(afterChain, 'From/to motion').frameRect.frameWidth - fromToMotionStart - 0.5) < 0.01, 'A from/to motion must end at its slide-percentage coordinate and accept percent-string attributes.')
  assert(Math.abs(shape(afterChain, 'From/by motion').frameRect.left / shape(afterChain, 'From/by motion').frameRect.frameWidth - fromByMotionStart - 0.35) < 0.01, 'A from/by motion must add the relative offset to its supplied start coordinate.')
  assert(Math.abs(shape(afterChain, 'To-only motion').frameRect.left / shape(afterChain, 'To-only motion').frameRect.frameWidth - toMotionStart - 0.4) < 0.01, 'A to-only motion must interpolate from the current layout position.')
  assert(Math.abs(shape(afterChain, 'Fixed motion path').frameRect.left / shape(afterChain, 'Fixed motion path').frameRect.frameWidth - 0.25) < 0.01 && fixedMotionStart > 0.3, 'A fixed layout-origin path must remain anchored to its slide coordinates when the target starts elsewhere.')
  const rotatedMotionEnd = shape(afterChain, 'Rotated fixed motion path').frameRect
  assert(Math.abs(rotatedMotionEnd.left / rotatedMotionEnd.frameWidth - 0.5) < 0.01 && Math.abs(rotatedMotionEnd.top / rotatedMotionEnd.frameHeight - (0.5 + 0.05 * rotatedMotionEnd.frameWidth / rotatedMotionEnd.frameHeight)) < 0.01 && rotatedMotionStart.left / rotatedMotionStart.frameWidth > 0.3, `rAng must rotate a fixed layout path clockwise around its rCtr slide point: ${JSON.stringify({ start: rotatedMotionStart, end: rotatedMotionEnd })}`)
  const formulaMotionEnd = shape(afterChain, 'Formula fixed motion path').frameRect
  assert(Math.abs(formulaMotionEnd.left / formulaMotionEnd.frameWidth - formulaMotionStart.left / formulaMotionStart.frameWidth - (2600000 / 12192000) * 1.5) < 0.01 && Math.abs(formulaMotionEnd.top / formulaMotionEnd.frameHeight - formulaMotionStart.top / formulaMotionStart.frameHeight - (2500000 / 6858000) * 0.5) < 0.01, `Static motion formulas must resolve target x/y/width/height and evaluate deterministic functions: ${JSON.stringify({ start: formulaMotionStart, end: formulaMotionEnd })}`)
  const parentMotionEnd = shape(afterChain, 'Parent origin grouped motion path').frameRect
  assert(Math.abs(parentMotionEnd.left / parentMotionEnd.frameWidth - parentMotionStart.left / parentMotionStart.frameWidth - 0.02) < 0.01 && Math.abs(parentMotionEnd.top / parentMotionEnd.frameHeight - parentMotionStart.top / parentMotionStart.frameHeight) < 0.01, `Parent-origin paths must apply the containing group's scale in slide coordinates: ${JSON.stringify({ start: parentMotionStart, end: parentMotionEnd })}`)
  const nestedParentMotionEnd = shape(afterChain, 'Nested group child').frameRect
  assert(Math.abs(nestedParentMotionEnd.left / nestedParentMotionEnd.frameWidth - nestedParentMotionStart.left / nestedParentMotionStart.frameWidth) < 0.01 && Math.abs(nestedParentMotionEnd.top / nestedParentMotionEnd.frameHeight - nestedParentMotionStart.top / nestedParentMotionStart.frameHeight + 0.01 * 12192000 * 4 / 6858000) < 0.01, `Parent-origin paths must apply nested group rotation, flip, and scale in slide coordinates: ${JSON.stringify({ start: nestedParentMotionStart, end: nestedParentMotionEnd })}`)
  assert(Math.abs(shape(afterChain, 'Curved motion path').frameRect.left / shape(afterChain, 'Curved motion path').frameRect.frameWidth - curveStart.left / curveStart.frameWidth) < 0.01, 'A closed multi-segment path must return to its starting x-coordinate.')
  assert(Math.abs(shape(afterChain, 'Curved motion path').frameRect.top / shape(afterChain, 'Curved motion path').frameRect.frameHeight - curveStart.top / curveStart.frameHeight) < 0.01, 'A closed multi-segment path must return to its starting y-coordinate.')
  assert(Math.abs(rotation(shape(afterChain, 'Scale emphasis'))) < 1, 'Rotation emphasis must remain dormant until its second-click step.')
  assert(shape(afterChain, 'Exit after previous').visibility === 'hidden' && shape(afterChain, 'Second click').visibility === 'hidden', 'The chained exit must finish before the next click step.')
  assert(shape(afterChain, 'Wipe exit').visibility === 'hidden' && shape(afterChain, 'Wipe exit').opacity === '1' && shape(afterChain, 'Wipe exit').wipeProgress === '0%', 'A wipe exit must close the mask and hide the target without fading it.')

  const scaleStartedAt = await evaluate(`(() => { const start = performance.now(); document.querySelector('.stage-shell').click(); return start })()`)
  await delay(30)
  const afterSecond = await state()
  assert(shape(afterSecond, 'Second click').visibility === 'visible', 'Second click must reveal its appear target.')
  assert(shape(afterSecond, 'Exit after previous').visibility === 'visible', 'A frozen exit must restore its base visibility when a later sibling animation starts.')
  assert(paragraph(afterSecond, 'Built paragraph').visibility === 'visible', 'The second click must reveal its targeted paragraph.')
  assert(rotation(shape(afterSecond, 'Scale emphasis')) > 0.5 && rotation(shape(afterSecond, 'Scale emphasis')) < 89.5, `A by-only rotation animation must interpolate its DrawingML angle: ${shape(afterSecond, 'Scale emphasis').transform}`)
  const scaleReverse = await stateAtElapsed(scaleStartedAt, 250)
  const scaleReverseRatio = transformScale(shape(scaleReverse, 'Scale emphasis'))
  assert(scaleReverseRatio > 1.15 && scaleReverseRatio < 1.45, 'Auto-reverse must interpolate back from the scale target before the next forward pass.')
  const scaleBeforeRemove = await stateAtElapsed(scaleStartedAt, 330)
  assert(transformScale(shape(scaleBeforeRemove, 'Scale emphasis')) > 1.01 && transformScale(shape(scaleBeforeRemove, 'Scale emphasis')) < 1.45, `The repeat duration must keep the scale animation active until its final phase: ${JSON.stringify(shape(scaleBeforeRemove, 'Scale emphasis'))}`)
  const scaleComplete = await stateAtElapsed(scaleStartedAt, 1_100)
  assert(Math.abs(transformScale(shape(scaleComplete, 'Scale emphasis')) - 1) < 0.02, 'A remove fill must restore the original scale after the repeat duration expires.')
  assert(Math.abs(rotation(shape(scaleComplete, 'Scale emphasis')) - 90) < 2, `A completed by-only rotation animation must hold its final angle: ${JSON.stringify({ angle: rotation(shape(scaleComplete, 'Scale emphasis')), transform: shape(scaleComplete, 'Scale emphasis').transform })}`)
  for (const direction of ['right', 'left', 'up', 'down']) {
    const wipe = shape(scaleComplete, `Wipe ${direction}`)
    const cssDirection = direction === 'up' ? 'top' : direction === 'down' ? 'bottom' : direction
    const directionMatches = cssDirection === 'bottom'
      ? !/linear-gradient\(to (top|left|right)/.test(wipe.maskImage)
      : wipe.maskImage.includes(`to ${cssDirection}`)
    assert(wipe.visibility === 'visible' && wipe.opacity === '1' && wipe.wipeProgress === '100%' && directionMatches, `The ${direction} wipe entrance must reveal with a mask in its OOXML direction and no fade: ${JSON.stringify(wipe)}`)
  }

  await evaluate(`document.querySelector('.player-controls button:first-child').click()`)
  await delay(30)
  const afterStepBack = await state()
  assert(shape(afterStepBack, 'Second click').visibility === 'hidden' && shape(afterStepBack, 'Exit after previous').visibility === 'hidden', 'Previous must undo only the latest click step.')
  assert(paragraph(afterStepBack, 'Built paragraph').visibility === 'hidden', 'Previous must undo the targeted paragraph build.')
  assert(Math.abs(shape(afterStepBack, 'Scale emphasis').frameRect.width / shape(initial, 'Scale emphasis').frameRect.width - 1) < 0.02, 'Stepping backward must restore the pre-emphasis scale.')

  await evaluate(`document.querySelector('.player-controls button:first-child').click()`)
  await delay(30)
  const afterAllBack = await state()
  assert(shape(afterAllBack, 'Click fade').visibility === 'hidden' && shape(afterAllBack, 'With previous').visibility === 'hidden', 'Undoing the first click must hide its entrances.')
  assert(Math.abs(shape(afterAllBack, 'Motion path').frameRect.left / shape(afterAllBack, 'Motion path').frameRect.frameWidth - motionStart) < 0.01, 'Undoing the click must restore the motion path start position.')
  assert(Math.abs(shape(afterAllBack, 'Curved motion path').frameRect.top / shape(afterAllBack, 'Curved motion path').frameRect.frameHeight - curveStart.top / curveStart.frameHeight) < 0.01, 'Undoing the click must restore the curved motion start position.')
  assert(shape(afterAllBack, 'Exit after previous').visibility === 'visible', 'Undoing the first click must restore its exit target.')

  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(30)
  const nextEventTriggered = await state()
  assert(shape(nextEventTriggered, 'Next event target').visibility === 'visible' && Number(shape(nextEventTriggered, 'Next event target').opacity) > 0, 'A slide-target onNext interactive sequence must run when advancing an animation step.')
  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(30)
  await evaluate(`document.querySelector('.player-controls button:first-child').click()`)
  await delay(30)
  const previousEventTriggered = await state()
  assert(shape(previousEventTriggered, 'Previous event target').visibility === 'visible' && Number(shape(previousEventTriggered, 'Previous event target').opacity) > 0, 'A slide-target onPrev interactive sequence must run when reversing an animation step.')
  await evaluate(`document.querySelector('.player-controls button:first-child').click()`)
  await delay(30)

  const speedStartedAt = await evaluate(`(() => { const started = performance.now(); document.querySelector('.stage-shell > .slide-host .slide-element[title="Fast speed trigger"]').click(); return started })()`)
  await delay(220)
  const fastAnimation = await state()
  assert(shape(fastAnimation, 'Fast speed fade').visibility === 'visible' && Number(shape(fastAnimation, 'Fast speed fade').opacity) > 0.98, `A 200% cTn speed must complete a 300ms effect in about 150ms (elapsed ${Math.round((await evaluate('performance.now()')) - speedStartedAt)}ms): ${JSON.stringify(shape(fastAnimation, 'Fast speed fade'))}`)

  const reverseBaseWidth = shape(initial, 'Reverse speed scale').frameRect.width
  const reverseSpeedStartedAt = await evaluate(`(() => { const started = performance.now(); document.querySelector('.stage-shell > .slide-host .slide-element[title="Reverse speed trigger"]').click(); return started })()`)
  await delay(50)
  const reverseMid = await state()
  const reverseMidScale = shape(reverseMid, 'Reverse speed scale').frameRect.width / reverseBaseWidth
  assert(reverseMidScale > 1.4 && reverseMidScale < 2.05, `A negative speed must start at the animation's target value and interpolate backward: ${JSON.stringify({ scale: reverseMidScale, shape: shape(reverseMid, 'Reverse speed scale') })}`)
  await delay(180)
  const reverseComplete = await state()
  const reverseEndScale = shape(reverseComplete, 'Reverse speed scale').frameRect.width / reverseBaseWidth
  assert(Math.abs(reverseEndScale - 1) < 0.02, `A negative-speed scale must finish at its from value after reversing the 300ms animation: ${JSON.stringify({ scale: reverseEndScale, elapsed: Math.round((await evaluate('performance.now()')) - reverseSpeedStartedAt) })}`)

  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(30)
  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(30)
  const fastForwarded = await state()
  assert(shape(fastForwarded, 'Click fade').opacity === '1' && shape(fastForwarded, 'Second click').visibility === 'visible', `A rapid next event must seek the active click step to its natural end before starting the next step: ${JSON.stringify({ clickFade: shape(fastForwarded, 'Click fade'), secondClick: shape(fastForwarded, 'Second click'), warnings: fastForwarded.warnings })}`)
  await evaluate(`document.querySelector('.player-controls button:first-child').click()`)
  await evaluate(`document.querySelector('.player-controls button:first-child').click()`)

  const paragraphFontSizeBase = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="Paragraph wipe target"] .text-run')).fontSize)`))
  const formulaFontSizeBase = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')).fontSize)`))
  const paragraphLetterSpacing = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="Paragraph wipe target"] .text-run')).letterSpacing)`))
  assert(Math.abs(paragraphLetterSpacing / paragraphFontSizeBase - 2.5 / 18) < 0.005, `DrawingML a:rPr/@spc must become letter spacing in hundredths of a point: ${paragraphLetterSpacing}/${paragraphFontSizeBase}`)
  const baselineShift = await evaluate(`(() => { const run = document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run'); return { offset: parseFloat(getComputedStyle(run).verticalAlign), size: parseFloat(getComputedStyle(run).fontSize) } })()`)
  assert(Math.abs(baselineShift.offset / baselineShift.size - 0.3) < 0.01, `DrawingML a:rPr/@baseline="30000" must raise the run by 30% of its font size: ${baselineShift.offset}/${baselineShift.size}`)
  const textShadowNames = ['Text shadow direct', 'Text shadow keyframes', 'Text shadow set none', 'Text shadow set auto']
  const dynamicMotionStart = await evaluate(`(() => { const target = document.querySelector('.stage-shell > .slide-host .slide-element[title="Dynamic formula motion path"]'); const frame = target.closest('.slide-frame').getBoundingClientRect(); const left = target.getBoundingClientRect().left - frame.left; const frameWidth = frame.width; const started = performance.now(); document.querySelector('.stage-shell > .slide-host .slide-element[title="Scale emphasis"]').click(); return { left, frameWidth, started } })()`)
  const triggerStartedAt = dynamicMotionStart.started
  const textEmbossStartFilters = await evaluate(`new Promise(resolve => requestAnimationFrame(() => resolve(${JSON.stringify(textShadowNames)}.map(name => getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="' + name + '"] .text-run')).filter))))`)
  await delay(80)
  const triggered = await state()
  const textRangeOpacityDuring = triggered.textRuns.filter(run => run.name === 'Text range opacity' && run.text.trim())
  const opacityFormulaTextRunsDuring = triggered.textRuns.filter(run => run.name === 'Opacity formula targets' && run.text.trim() === 'Run')
  const initialTextShadowRuns = initial.textRuns.filter(run => textShadowNames.includes(run.name) && run.text.trim())
  const textShadowRuns = (snapshot, name) => snapshot.textRuns.filter(run => run.name === name && run.text.trim())
  const textEmbossFilter = name => evaluate(`getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="${name}"] .text-run')).filter`)
  assert(initialTextShadowRuns.length === textShadowNames.length && initialTextShadowRuns.every(run => run.textShadow !== 'none'), `Text-shadow animation fixtures must begin with a static run shadow: ${JSON.stringify(initialTextShadowRuns)}`)
  assert(textShadowRuns(triggered, 'Text shadow direct').every(run => run.textShadow !== 'none') && textShadowRuns(triggered, 'Text shadow keyframes').every(run => run.textShadow !== 'none') && textShadowRuns(triggered, 'Text shadow set none').every(run => run.textShadow === 'none') && textShadowRuns(triggered, 'Text shadow set auto').every(run => run.textShadow !== 'none'), `Direct/from and p:set text-shadow values must compose with the static shadow at trigger start: ${JSON.stringify(textShadowNames.map(name => [name, textShadowRuns(triggered, name)]))}`)
  assert(textEmbossStartFilters[0] === 'none' && textEmbossStartFilters[1] === 'none' && textEmbossStartFilters[2]?.includes('drop-shadow') && textEmbossStartFilters[3] === 'none', `Text emboss from/to, keyframes, and p:set must apply discrete values: ${JSON.stringify(textEmbossStartFilters)}`)
  assert(triggered.warnings.some(warning => warning.includes('text emboss animation is approximated with CSS drop shadows')), `Text emboss approximation must be reported: ${JSON.stringify(triggered.warnings)}`)
  assert(triggered.warnings.some(warning => warning.includes('text shadow animation without complete static shadows or with a text range was skipped.')), `Unsupported shadow ranges and missing static shadows must produce a warning: ${JSON.stringify(triggered.warnings)}`)
  const cssAlpha = value => Number(value.match(/rgba\([^,]+,\s*[^,]+,\s*[^,]+,\s*([\d.]+)\)/)?.[1] ?? 1)
  assert(textRangeOpacityDuring.length === 2 && [textRangeOpacityDuring[0].color, textRangeOpacityDuring[0].textStrokeColor, textRangeOpacityDuring[0].textShadow, textRangeOpacityDuring[1].color, textRangeOpacityDuring[1].textStrokeColor, textRangeOpacityDuring[1].textShadow].every(value => cssAlpha(value) > 0 && cssAlpha(value) < 1), 'Numeric and formula character-range fill, outline, and shadow opacity must interpolate only across their selected text.')
  assert(opacityFormulaTextRunsDuring.length === 1 && opacityFormulaTextRunsDuring[0].opacity > 0 && opacityFormulaTextRunsDuring[0].opacity < 1, `A style.opacity formula must apply to its targeted text range: ${JSON.stringify(opacityFormulaTextRunsDuring)}`)
  const cssRgb = value => value.match(/rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)/)?.slice(1, 4).map(Number)
  const textRangeShadowDuringRgb = cssRgb(textRangeOpacityDuring[0]?.textShadow || '')
  assert(textRangeShadowDuringRgb?.[0] > 0 && textRangeShadowDuringRgb[0] < 255 && textRangeShadowDuringRgb[1] === 0 && textRangeShadowDuringRgb[2] === 0, `Character-range shadow.color must interpolate on only the selected text: ${textRangeOpacityDuring[0]?.textShadow}`)
  assert(triggered.slide?.startsWith('01 / 67') && shape(triggered, 'Click fade').visibility === 'visible', 'Clicking the configured trigger shape must start its animation without advancing the slide build.')
  assert(shape(triggered, 'Dynamic formula motion path').frameRect.left < dynamicMotionStart.left - dynamicMotionStart.frameWidth * 0.04, `A dynamic motion formula must change the target path as progress advances: ${shape(triggered, 'Dynamic formula motion path').frameRect.left - dynamicMotionStart.left}`)
  const triggerOpacity = Number(shape(triggered, 'Click fade').opacity)
  const formulaOpacityDuring = Number(shape(triggered, 'By word').opacity)
  assert(formulaOpacityDuring > 0.4 && formulaOpacityDuring < 0.8, `A linear style.opacity p:tav formula must affect the target element: ${formulaOpacityDuring}`)
  const formulaFontSizeDuring = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')).fontSize)`)) / formulaFontSizeBase
  assert(formulaFontSizeDuring > 1.3 && formulaFontSizeDuring < 1.8, `A linear style.fontSize p:tav formula must affect the active text size: ${formulaFontSizeDuring}`)
  const fontWeightDuring = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="Paragraph wipe target"] .text-run')).fontWeight)`))
  const keyframedFontWeightSample = await evaluate(`(() => ({ weight: parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By letter"] .text-run')).fontWeight), now: performance.now() }))()`)
  const keyframedFontWeightDuring = Number(keyframedFontWeightSample.weight)
  assert(fontWeightDuring > 400 && fontWeightDuring < 900, `A generic style.fontWeight animation must interpolate text weight: ${fontWeightDuring}`)
  const keyframedFontWeightProgress = Math.min(1, Math.max(0, (Number(keyframedFontWeightSample.now) - triggerStartedAt) / 300))
  const expectedKeyframedFontWeight = keyframedFontWeightProgress < 0.5 ? 400 + keyframedFontWeightProgress * 600 : 700 + (keyframedFontWeightProgress - 0.5) * 400
  assert(Math.abs(keyframedFontWeightDuring - expectedKeyframedFontWeight) < 50, `A numeric font-weight p:tavLst must interpolate from its keyframes at the captured animation time: ${keyframedFontWeightDuring}/${expectedKeyframedFontWeight}`)
  assert(shape(triggered, 'Custom geometry').opacity === '0.5', `A triggered p:set fltVal must apply its fixed opacity: ${shape(triggered, 'Custom geometry').opacity}`)
  assert(shape(triggered, 'Custom geometry').customPaths[0]?.fill === 'rgb(255, 0, 0)', `A triggered p:set clrVal must apply its fixed fill color: ${JSON.stringify(shape(triggered, 'Custom geometry').customPaths)}`)
  assert(shape(triggered, 'Paragraph wipe target').opacity === '0.75', `A triggered p:set strVal must apply its numeric opacity: ${shape(triggered, 'Paragraph wipe target').opacity}`)
  assert(triggerOpacity > 0 && triggerOpacity < 1, 'A shape-triggered fade must interpolate from the trigger click.')
  const opacityDuring = Number(shape(triggered, 'Fast speed fade').opacity)
  assert(opacityDuring > 0.6 && opacityDuring < 1, `A generic style.opacity animation must interpolate on its target: ${opacityDuring}`)
  const keyframedOpacityDuring = Number(shape(triggered, 'Stroke color').opacity)
  assert(keyframedOpacityDuring > 0.2 && keyframedOpacityDuring < 0.8, `A generic opacity p:tavLst must interpolate between its numeric keypoints: ${keyframedOpacityDuring}`)
  assert(Math.abs(Number(shape(triggered, 'Scale emphasis').opacity) - 0.3) < 0.02, `A discrete opacity p:tavLst must hold its initial keypoint: ${shape(triggered, 'Scale emphasis').opacity}`)
  const fontSizeStartRatio = Number.parseFloat(triggered.textRuns.find(run => run.name === 'Paragraph wipe target' && run.text.trim())?.fontSize || '') / paragraphFontSizeBase
  assert(Math.abs(fontSizeStartRatio - 0.5) < 0.05, `A discrete style.fontSize p:tavLst must hold its initial keypoint: ${fontSizeStartRatio}`)
  const fillAlphaDuring = Number(/rgba\([^)]*,\s*([\d.]+)\)$/.exec(shape(triggered, 'With previous').backgroundColor)?.[1])
  assert(fillAlphaDuring > 0.5 && fillAlphaDuring < 1, `A generic fill.opacity animation must interpolate only the shape fill: ${shape(triggered, 'With previous').backgroundColor}`)
  const strokeAlphaDuring = Number(/rgba\([^)]*,\s*([\d.]+)\)$/.exec(shape(triggered, 'Stroke color').lineColor)?.[1])
  assert(strokeAlphaDuring > 0.5 && strokeAlphaDuring < 1, `A generic stroke.opacity animation must interpolate only the line stroke: ${shape(triggered, 'Stroke color').lineColor}`)
  const shadowAlphaDuring = Number(/rgba\([^)]*,\s*([\d.]+)\)/.exec(shape(triggered, 'With previous').boxShadow)?.[1])
  assert(shadowAlphaDuring > 0.25 && shadowAlphaDuring < 0.5, `A generic shadow.opacity animation must interpolate only shadow alpha: ${shape(triggered, 'With previous').boxShadow}`)
  const shadowRedDuring = Number(/rgba\(([\d.]+)/.exec(shape(triggered, 'With previous').boxShadow)?.[1])
  assert(shadowRedDuring > 0 && shadowRedDuring < 255, `A shadow.color animation must interpolate its target color: ${shape(triggered, 'With previous').boxShadow}`)
  assert(shape(triggered, 'Rotated group child').transform !== shape(initial, 'Rotated group child').transform, 'Rotation animation must compose with a grouped shape’s existing transform matrix.')
  const hueSample = await evaluate(`new Promise(resolve => {
    const sample = () => {
      const color = getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="Stroke color"] .line-render')).borderTopColor
      const [red, green, blue] = (color.match(/\\d+/g) || []).map(Number)
      if (red > 150 && blue > 150 && green < 60 || performance.now() - ${triggerStartedAt} >= 300) resolve({ color, elapsed: performance.now() - ${triggerStartedAt}, passed: red > 150 && blue > 150 && green < 60 })
      else requestAnimationFrame(sample)
    }
    sample()
  })`)
  assert(hueSample.passed, `Clockwise HSL interpolation must pass through magenta: ${JSON.stringify(hueSample)}`)
  await delay(Math.max(0, 190 - ((await evaluate('performance.now()')) - triggerStartedAt)))
  const discreteOpacityMiddle = await state()
  assert(Math.abs(Number(shape(discreteOpacityMiddle, 'Scale emphasis').opacity) - 0.8) < 0.02, `A discrete opacity p:tavLst must jump at its middle keypoint: ${shape(discreteOpacityMiddle, 'Scale emphasis').opacity}`)
  const iterationMiddle = await state()
  const runsNamed = name => iterationMiddle.textRuns.filter(run => run.name === name && run.text.trim())
  const paragraphWipeMiddle = runsNamed('Paragraph wipe target')[0]
  const byLetterMiddle = runsNamed('By letter')[0]
  const byWordMiddleRuns = runsNamed('By word')
  const opacityWordMiddleRuns = runsNamed('Text range opacity')
  const fontSizeMiddleRatio = Number.parseFloat(paragraphWipeMiddle?.fontSize || '') / paragraphFontSizeBase
  assert(Math.abs(fontSizeMiddleRatio - 2) < 0.05, `A discrete style.fontSize p:tavLst must jump at its middle keypoint: ${fontSizeMiddleRatio}`)
  const fontWeightMiddle = Number.parseFloat(paragraphWipeMiddle?.fontWeight || '')
  const keyframedFontWeightMiddle = Number.parseFloat(byLetterMiddle?.fontWeight || '')
  const formulaFontWeightMiddle = Number.parseFloat(byWordMiddleRuns[0]?.fontWeight || '')
  const formulaFontSizeMiddle = Number.parseFloat(byWordMiddleRuns[0]?.fontSize || '') / formulaFontSizeBase
  const formulaOpacityMiddle = Number(shape(iterationMiddle, 'By word').opacity)
  const opacityFormulaMidpoint = shape(iterationMiddle, 'Opacity formula targets')
  const opacityFormulaFillMiddle = cssAlpha(opacityFormulaMidpoint.backgroundColor)
  const opacityFormulaStrokeMiddle = cssAlpha(opacityFormulaMidpoint.borderColor)
  const opacityFormulaShadowMiddle = cssAlpha(opacityFormulaMidpoint.boxShadow)
  const textStyleFontStyleMiddle = byWordMiddleRuns[0]?.fontStyle
  const textStyleUnderlineMiddle = byWordMiddleRuns[0]?.textDecorationLine
  const textStyleLineThroughMiddle = byWordMiddleRuns.map(run => run.textDecorationLine)
  const directLineThroughMiddle = iterationMiddle.textRuns.filter(run => run.name === 'Text range opacity' && run.text.trim())
  const fontFamilyKeyframesMiddle = byWordMiddleRuns.map(run => run.fontFamily.toLowerCase())
  const textTransformMiddle = byWordMiddleRuns[0]?.verticalAlign
  const textOutlineMiddle = Number.parseFloat(byWordMiddleRuns[0]?.textStrokeWidth || '')
  const textShadowKeyframeMiddle = runsNamed('Text shadow keyframes')
  const textEmbossKeyframeMiddle = await textEmbossFilter('Text shadow keyframes')
  const iteratedWordMiddle = byWordMiddleRuns.map(run => ({ text: run.text, weight: Number.parseFloat(run.fontWeight), fontStyle: run.fontStyle }))
  const iteratedOpacityMiddle = opacityWordMiddleRuns.map(run => ({ text: run.text, opacity: Number.parseFloat(run.opacity) }))
  assert(fontWeightMiddle > 700 && fontWeightMiddle < 900, `A held font-weight value must progress during its second half: ${fontWeightMiddle}`)
  assert(keyframedFontWeightMiddle > 700 && keyframedFontWeightMiddle < 900, `A linear numeric font-weight p:tavLst must interpolate the final interval: ${keyframedFontWeightMiddle}`)
  assert(formulaFontWeightMiddle > 700 && formulaFontWeightMiddle < 850, `A linear style.fontWeight p:tav formula must combine the numeric track and behavior progress: ${formulaFontWeightMiddle}`)
  assert(formulaFontSizeMiddle > 1.6 && formulaFontSizeMiddle < 2, `A linear style.fontSize p:tav formula must combine the numeric track and behavior progress: ${formulaFontSizeMiddle}`)
  assert(formulaOpacityMiddle > 0.6 && formulaOpacityMiddle < 0.9, `A linear style.opacity p:tav formula must combine the numeric track and behavior progress: ${formulaOpacityMiddle}`)
  assert(opacityFormulaFillMiddle > 0.55 && opacityFormulaFillMiddle < 0.8 && opacityFormulaStrokeMiddle > 0.45 && opacityFormulaStrokeMiddle < 0.7 && opacityFormulaShadowMiddle > 0.15 && opacityFormulaShadowMiddle < 0.3, `Linear p:tav formulas must animate fill.opacity, stroke.opacity, and shadow.opacity: ${JSON.stringify({ opacityFormulaFillMiddle, opacityFormulaStrokeMiddle, opacityFormulaShadowMiddle, opacityFormulaMidpoint })}`)
  assert(textStyleFontStyleMiddle === 'normal' && textStyleUnderlineMiddle.includes('underline') && textStyleLineThroughMiddle[0]?.includes('line-through') && !textStyleLineThroughMiddle[1]?.includes('line-through') && textTransformMiddle === 'baseline', `Discrete text-style keyframes must compose underline and stagger line-through by word: ${textStyleFontStyleMiddle}/${JSON.stringify(textStyleLineThroughMiddle)}/${textTransformMiddle}`)
  assert(directLineThroughMiddle.length > 0 && directLineThroughMiddle.every(run => !run.textDecorationLine.includes('line-through')), `A discrete line-through from/to animation must apply its false value during playback: ${JSON.stringify(directLineThroughMiddle)}`)
  assert(fontFamilyKeyframesMiddle.length > 0 && fontFamilyKeyframesMiddle.every(fontFamily => fontFamily.includes('georgia')), `A discrete style.fontFamily p:tavLst must apply its active string keyframe: ${JSON.stringify(fontFamilyKeyframesMiddle)}`)
  assert(textOutlineMiddle === 0, `A discrete text outline animation must preserve its false value before completion: ${textOutlineMiddle}`)
  assert(textShadowRuns(iterationMiddle, 'Text shadow direct').every(run => run.textShadow !== 'none') && textShadowKeyframeMiddle.length > 0 && textShadowKeyframeMiddle.every(run => run.textShadow === 'none'), `Discrete text-shadow keyframes must hold their middle none value while a from/to animation holds normal: ${JSON.stringify({ direct: textShadowRuns(iterationMiddle, 'Text shadow direct'), keyframes: textShadowKeyframeMiddle })}`)
  assert(textEmbossKeyframeMiddle.includes('drop-shadow'), `The textEffectEmboss middle keyframe must enable its embossed filter: ${textEmbossKeyframeMiddle}`)
  assert(iteratedWordMiddle.length === 2 && iteratedWordMiddle[0].weight > iteratedWordMiddle[1].weight && iteratedWordMiddle[1].weight > 400 && iteratedWordMiddle.every(run => run.fontStyle === 'normal'), `p:iterate by word must stagger font-weight and discrete text style: ${JSON.stringify(iteratedWordMiddle)}`)
  assert(iteratedOpacityMiddle.length === 2 && iteratedOpacityMiddle[0].opacity > iteratedOpacityMiddle[1].opacity && iteratedOpacityMiddle[1].opacity > 0.2 && iteratedOpacityMiddle[0].opacity < 1, `p:iterate by word must stagger generic style.opacity on text runs: ${JSON.stringify(iteratedOpacityMiddle)}`)
  await delay(130)
  const triggerComplete = await state()
  const textRangeOpacityFinal = triggerComplete.textRuns.filter(run => run.name === 'Text range opacity' && run.text.trim())
  const opacityFormulaTextRunsFinal = triggerComplete.textRuns.filter(run => run.name === 'Opacity formula targets' && run.text.trim() === 'Run')
  const textEmbossFinalFilters = await Promise.all(textShadowNames.map(name => textEmbossFilter(name)))
  assert(textEmbossFinalFilters[0]?.includes('drop-shadow') && textEmbossFinalFilters[1] === 'none' && textEmbossFinalFilters[2]?.includes('drop-shadow') && textEmbossFinalFilters[3] === 'none', `Text emboss animations and p:set must hold or reset their final value: ${JSON.stringify(textEmbossFinalFilters)}`)
  assert(textShadowRuns(triggerComplete, 'Text shadow direct').every(run => run.textShadow === 'none') && textShadowRuns(triggerComplete, 'Text shadow keyframes').every(run => run.textShadow !== 'none') && textShadowRuns(triggerComplete, 'Text shadow set none').every(run => run.textShadow === 'none') && textShadowRuns(triggerComplete, 'Text shadow set auto').every(run => run.textShadow !== 'none'), `Discrete text-shadow animations and p:set must hold none or restore the static shadow for auto/normal: ${JSON.stringify(textShadowNames.map(name => [name, textShadowRuns(triggerComplete, name)]))}`)
  const iteratedFontFamilyFinal = triggerComplete.textRuns.filter(run => run.name === 'Text range opacity' && run.text.trim())
  const byLetterFontFamily = triggerComplete.textRuns.filter(run => run.name === 'By letter' && run.text.trim())
  const unsafeFontFamilyRuns = triggerComplete.textRuns.filter(run => run.name === 'Paragraph wipe target' && run.text.trim())
  assert(iteratedFontFamilyFinal.length === 2 && iteratedFontFamilyFinal[0].fontFamily.toLowerCase().includes('courier new') && iteratedFontFamilyFinal[1].fontFamily.toLowerCase().includes('arial'), `p:iterate by word must finish the first discrete font-family value while the second word remains at its start value: ${JSON.stringify(iteratedFontFamilyFinal)}`)
  assert(byLetterFontFamily.length > 0 && byLetterFontFamily.every(run => run.fontFamily.toLowerCase().includes('georgia')), `A p:set style.fontFamily must apply to the text frame: ${JSON.stringify(byLetterFontFamily)}`)
  assert(unsafeFontFamilyRuns.length > 0 && unsafeFontFamilyRuns.every(run => !run.fontFamily.toLowerCase().includes('color:red')), `An unsafe font-family value must not reach rendered CSS: ${JSON.stringify(unsafeFontFamilyRuns)}`)
  assert(textRangeOpacityFinal.length === 2 && Math.abs(cssAlpha(textRangeOpacityFinal[0].color) - 0.25) < 0.03 && Math.abs(cssAlpha(textRangeOpacityFinal[0].textStrokeColor) - 0.25) < 0.03 && Math.abs(cssAlpha(textRangeOpacityFinal[0].textShadow) - 0.25) < 0.03 && Math.abs(cssAlpha(textRangeOpacityFinal[1].color) - 0.4) < 0.03 && Math.abs(cssAlpha(textRangeOpacityFinal[1].textStrokeColor) - 0.5) < 0.03 && Math.abs(cssAlpha(textRangeOpacityFinal[1].textShadow) - 0.4) < 0.03, `Numeric and formula range opacity animations must hold their final alpha only on the targeted text: ${JSON.stringify(textRangeOpacityFinal)}`)
  assert(opacityFormulaTextRunsFinal.length === 1 && Math.abs(opacityFormulaTextRunsFinal[0].opacity - 0.4) < 0.03, `A style.opacity formula must hold its final value on the targeted text range: ${JSON.stringify(opacityFormulaTextRunsFinal)}`)
  const textRangeShadowFinalRgb = cssRgb(textRangeOpacityFinal[0]?.textShadow || '')
  const unselectedShadowFinalRgb = cssRgb(textRangeOpacityFinal[1]?.textShadow || '')
  assert(textRangeShadowFinalRgb?.[0] === 255 && textRangeShadowFinalRgb[1] === 0 && textRangeShadowFinalRgb[2] === 0 && unselectedShadowFinalRgb?.[0] === 0 && unselectedShadowFinalRgb[1] === 0 && unselectedShadowFinalRgb[2] === 0, `Character-range shadow.color must hold red only on selected text: ${JSON.stringify(textRangeOpacityFinal)}`)
  assert(Number(shape(triggerComplete, 'Click fade').opacity) === 1, 'A shape-triggered animation must hold its final visual state.')
  const fontWeightEnd = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="Paragraph wipe target"] .text-run')).fontWeight)`))
  const keyframedFontWeightEnd = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By letter"] .text-run')).fontWeight)`))
  const formulaFontWeightEnd = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')).fontWeight)`))
  const formulaFontSizeEnd = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')).fontSize)`)) / formulaFontSizeBase
  const iteratedWordEnd = await evaluate(`[...document.querySelectorAll('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')].filter(run => run.textContent.trim()).map(run => ({ text: run.textContent, weight: Number(getComputedStyle(run).fontWeight), fontStyle: getComputedStyle(run).fontStyle }))`)
  const iteratedOpacityEnd = textRangeOpacityFinal.map(run => ({ text: run.text, opacity: Number(run.opacity) }))
  const formulaOpacityEnd = Number(shape(triggerComplete, 'By word').opacity)
  const textStyleFontStyleEnd = await evaluate(`getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')).fontStyle`)
  const textStyleUnderlineEnd = await evaluate(`getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')).textDecorationLine`)
  const directLineThroughEnd = await evaluate(`getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="Text range opacity"] .text-run')).textDecorationLine`)
  const pSetFontStyle = await evaluate(`getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By letter"] .text-run')).fontStyle`)
  const pSetUnderline = await evaluate(`getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By letter"] .text-run')).textDecorationLine`)
  const textTransformEnd = await evaluate(`getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')).verticalAlign`)
  const pSetTextTransform = await evaluate(`getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By letter"] .text-run')).verticalAlign`)
  const textOutlineEnd = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')).webkitTextStrokeWidth)`))
  const pSetTextOutline = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="Paragraph wipe target"] .text-run')).webkitTextStrokeWidth)`))
  assert(fontWeightEnd === 900 && keyframedFontWeightEnd === 900 && formulaFontWeightEnd === 700, `Font-weight animations must hold their final values: ${fontWeightEnd}/${keyframedFontWeightEnd}/${formulaFontWeightEnd}`)
  assert(Math.abs(formulaFontSizeEnd - 1.5) < 0.03, `A linear style.fontSize p:tav formula must hold its final numeric keyframe: ${formulaFontSizeEnd}`)
  assert(Math.abs(formulaOpacityEnd - 0.7) < 0.03, `A linear style.opacity p:tav formula must hold its final numeric keyframe: ${formulaOpacityEnd}`)
  assert(textStyleFontStyleEnd === 'italic' && textStyleUnderlineEnd.includes('underline') && textStyleUnderlineEnd.includes('line-through') && directLineThroughEnd.includes('line-through'), `Discrete font-style, underline, and line-through animations must hold their final values: ${textStyleFontStyleEnd}/${textStyleUnderlineEnd}/${directLineThroughEnd}`)
  assert(iteratedWordEnd.length === 2 && iteratedWordEnd[0].weight === 700 && iteratedWordEnd[1].weight > 700 && iteratedWordEnd[1].weight < 850 && iteratedWordEnd[0].fontStyle === 'italic' && iteratedWordEnd[1].fontStyle === 'normal', `The first iterated word must finish before the second word begins its final state: ${JSON.stringify(iteratedWordEnd)}`)
  assert(iteratedOpacityEnd.length === 2 && iteratedOpacityEnd[0].opacity === 1 && iteratedOpacityEnd[1].opacity > 0.5 && iteratedOpacityEnd[1].opacity < 0.9, `p:iterate opacity must hold the first word while animating the second: ${JSON.stringify(iteratedOpacityEnd)}`)
  assert(pSetFontStyle === 'italic' && pSetUnderline.includes('underline') && pSetUnderline.includes('line-through'), `p:set must compose its discrete font-style, underline, and line-through values: ${pSetFontStyle}/${pSetUnderline}`)
  assert(textTransformEnd === 'super' && pSetTextTransform === 'sub', `Discrete p:anim and p:set textTransform values must map to superscript and subscript: ${textTransformEnd}/${pSetTextTransform}`)
  assert(textOutlineEnd > 0 && pSetTextOutline === 0, `Discrete textEffectOutline animations must enable parsed outlines and let p:set hide them: ${textOutlineEnd}/${pSetTextOutline}`)
  await delay(150)
  const iteratedWordComplete = await evaluate(`[...document.querySelectorAll('.stage-shell > .slide-host .slide-element[title="By word"] .text-run')].filter(run => run.textContent.trim()).map(run => ({ text: run.textContent, weight: Number(getComputedStyle(run).fontWeight), fontStyle: getComputedStyle(run).fontStyle, fontFamily: getComputedStyle(run).fontFamily.toLowerCase() }))`)
  const iteratedOpacityComplete = await evaluate(`[...document.querySelectorAll('.stage-shell > .slide-host .slide-element[title="Text range opacity"] .text-run')].filter(run => run.textContent.trim()).map(run => Number(getComputedStyle(run).opacity))`)
  const iteratedFontFamilyComplete = await evaluate(`[...document.querySelectorAll('.stage-shell > .slide-host .slide-element[title="Text range opacity"] .text-run')].filter(run => run.textContent.trim()).map(run => getComputedStyle(run).fontFamily.toLowerCase())`)
  assert(iteratedWordComplete.length === 2 && iteratedWordComplete.every(run => run.weight === 700 && run.fontStyle === 'italic'), `p:iterate by word must hold the final numeric and discrete text-style values for every word: ${JSON.stringify(iteratedWordComplete)}`)
  assert(iteratedWordComplete.every(run => run.fontFamily.includes('courier new')), `A discrete style.fontFamily p:tavLst must hold its final string keyframe: ${JSON.stringify(iteratedWordComplete)}`)
  assert(iteratedOpacityComplete.length === 2 && iteratedOpacityComplete.every(opacity => opacity === 1), `p:iterate opacity must hold the final value for every word: ${JSON.stringify(iteratedOpacityComplete)}`)
  assert(iteratedFontFamilyComplete.length === 2 && iteratedFontFamilyComplete.every(fontFamily => fontFamily.includes('courier new')), `p:iterate by word must hold the final font-family value for every word: ${JSON.stringify(iteratedFontFamilyComplete)}`)
  assert(Math.abs(Number(shape(triggerComplete, 'Fast speed fade').opacity) - 0.25) < 0.02, `A generic style.opacity animation must hold its final value: ${shape(triggerComplete, 'Fast speed fade').opacity}`)
  assert(Math.abs(Number(shape(triggerComplete, 'Stroke color').opacity) - 0.4) < 0.02, `A generic opacity p:tavLst must hold its last keypoint: ${shape(triggerComplete, 'Stroke color').opacity}`)
  assert(Math.abs(Number(shape(triggerComplete, 'Scale emphasis').opacity) - 0.4) < 0.02, `A discrete opacity p:tavLst must hold its last keypoint: ${shape(triggerComplete, 'Scale emphasis').opacity}`)
  const fontSizeEndRatio = Number(await evaluate(`parseFloat(getComputedStyle(document.querySelector('.stage-shell > .slide-host .slide-element[title="Paragraph wipe target"] .text-run')).fontSize)`)) / paragraphFontSizeBase
  assert(Math.abs(fontSizeEndRatio - 0.75) < 0.05, `A discrete style.fontSize p:tavLst must hold its last keypoint: ${fontSizeEndRatio}`)
  assert(shape(triggerComplete, 'With previous').backgroundColor === 'rgba(255, 0, 0, 0.5)', `A generic fill.opacity animation must hold its final value over the animated fill: ${shape(triggerComplete, 'With previous').backgroundColor}`)
  assert(shape(triggerComplete, 'Stroke color').lineColor === 'rgba(255, 0, 0, 0.5)', `A generic stroke.opacity animation must hold its final value over the animated stroke: ${shape(triggerComplete, 'Stroke color').lineColor}`)
  assert(shape(triggerComplete, 'With previous').boxShadow.includes('rgba(255, 0, 0, 0.25)'), `A generic shadow.opacity animation must hold its final alpha: ${shape(triggerComplete, 'With previous').boxShadow}`)
  const opacityFormulaEnd = shape(triggerComplete, 'Opacity formula targets')
  assert(Math.abs(cssAlpha(opacityFormulaEnd.backgroundColor) - 0.5) < 0.03 && Math.abs(cssAlpha(opacityFormulaEnd.borderColor) - 0.4) < 0.03 && Math.abs(cssAlpha(opacityFormulaEnd.boxShadow) - 0.15) < 0.03, `Opacity formula animations must hold their final keyframe values: ${JSON.stringify(opacityFormulaEnd)}`)
  assert(shape(triggerComplete, 'Rotated group child').transform !== shape(initial, 'Rotated group child').transform, 'A completed grouped-shape rotation must hold its transformed matrix.')
  const dynamicMotionEnd = await stateAtElapsed(triggerStartedAt, 800)
  assert(Math.abs(shape(dynamicMotionEnd, 'Dynamic formula motion path').frameRect.left - dynamicMotionStart.left) < dynamicMotionStart.frameWidth * 0.01, `A sampled dynamic motion formula must reach its end-time path position: ${shape(dynamicMotionEnd, 'Dynamic formula motion path').frameRect.left - dynamicMotionStart.left}`)
  assert(shape(triggerComplete, 'With previous').backgroundColor === 'rgba(255, 0, 0, 0.5)', 'A triggered fillcolor animation must hold its final fill under fill opacity.')
  assert(shape(triggerComplete, 'Stroke color').lineColor === 'rgba(255, 0, 0, 0.5)', 'A triggered stroke.color animation must hold its final line color under stroke opacity.')
  assert(paragraph(triggerComplete, 'Built paragraph').color === 'rgb(255, 0, 255)' && paragraph(triggerComplete, 'Al\nways visible').color !== 'rgb(255, 0, 255)', 'A paragraph-range style.color animation must affect only its targeted paragraph.')
  assert(shape(triggerComplete, 'With previous').visibility === 'hidden', 'A shape-trigger click must not advance the main click-build sequence.')
  const slideIn = shape(triggered, 'Slide entrance and exit')
  const slideInX = transformValues(slideIn)[4]
  assert(slideIn.visibility === 'visible' && slideInX < -20 && slideInX > -slideIn.frameRect.frameWidth * 1.1, `slide(fromLeft) must move the target into view from the slide's left edge: ${JSON.stringify(slideIn)}`)
  await evaluate(`document.querySelector('.stage-shell > .slide-host .slide-element[title="Scale emphasis"]').click()`)
  await delay(40)
  const slideOut = shape(await state(), 'Slide entrance and exit')
  assert(slideOut.visibility === 'visible' && transformValues(slideOut)[4] > 0, `slide(fromRight) exit must move the target toward the right edge: ${JSON.stringify(slideOut)}`)
  await delay(300)
  const slideOutComplete = shape(await state(), 'Slide entrance and exit')
  assert(slideOutComplete.visibility === 'hidden' && transformValues(slideOutComplete)[4] > slideOutComplete.frameRect.frameWidth * 0.9, `A held slide exit must finish beyond the right edge: ${JSON.stringify(slideOutComplete)}`)
  await evaluate(`document.querySelector('.stage-shell > .slide-host .slide-element[title="Scale emphasis"]').dispatchEvent(new MouseEvent('mouseenter'))`)
  await delay(250)
  assert(shape(await state(), 'Second click').visibility === 'visible', 'An onMouseOver interactive sequence must run when its trigger shape is entered.')
  await evaluate(`document.querySelector('.stage-shell > .slide-host .slide-element[title="Scale emphasis"]').dispatchEvent(new MouseEvent('mouseleave'))`)
  await delay(250)
  assert(shape(await state(), 'Second click').visibility === 'hidden', 'An onMouseOut interactive sequence must run when its trigger shape is left.')
  await evaluate(`document.querySelector('.stage-shell > .slide-host .slide-element[title="Nested group child"]').dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))`)
  await delay(250)
  assert(shape(await state(), 'Second click').visibility === 'visible', 'An onDblClick interactive sequence must run on its trigger shape.')

  const repeatCountStartedAt = await evaluate(`(() => { const started = performance.now(); document.querySelector('.stage-shell > .slide-host .slide-element[title="Parent repeat count trigger"]').click(); return started })()`)
  await delay(660)
  const repeatedParent = await state()
  const repeatCountOpacity = Number(shape(repeatedParent, 'Parent repeat count target').opacity)
  assert(repeatCountOpacity > 0.05 && repeatCountOpacity < 0.95, `A finite parent repeatCount must restart the child fade on its second cycle: ${JSON.stringify({ elapsed: repeatedParent.now - repeatCountStartedAt, opacity: repeatCountOpacity, warnings: repeatedParent.warnings })}`)

  const repeatDurationStartedAt = await evaluate(`(() => { const started = performance.now(); document.querySelector('.stage-shell > .slide-host .slide-element[title="Parent repeat duration trigger"]').click(); return started })()`)
  await delay(950)
  const repeatedParentDuration = await state()
  const afterRepeatOpacity = Number(shape(repeatedParentDuration, 'Parent repeat after-effect').opacity)
  assert(Number(shape(repeatedParentDuration, 'Parent repeat duration target').opacity) === 1 && afterRepeatOpacity > 0.05 && afterRepeatOpacity < 0.95, `A finite parent repeatDur must cap looping and defer the after-effect until the cap: ${JSON.stringify({ elapsed: repeatedParentDuration.now - repeatDurationStartedAt, repeatedTarget: shape(repeatedParentDuration, 'Parent repeat duration target'), afterEffect: shape(repeatedParentDuration, 'Parent repeat after-effect'), warnings: repeatedParentDuration.warnings })}`)

  const reverseParent = await evaluate(`new Promise(resolve => {
    const trigger = document.querySelector('.stage-shell > .slide-host .slide-element[title="Parent reverse speed trigger"]')
    const target = document.querySelector('.stage-shell > .slide-host .slide-element[title="Parent reverse speed fade"]')
    if (!trigger || !target) return resolve({ passed: false, reason: 'fixture shape missing' })
    let sawEndState = false
    const startedAt = performance.now()
    trigger.click()
    const check = () => {
      const style = getComputedStyle(target)
      const opacity = Number(style.opacity)
      const visible = style.visibility !== 'hidden'
      if (visible && opacity > 0.8) sawEndState = true
      if (sawEndState && visible && opacity < 0.1) return resolve({ passed: true, elapsed: performance.now() - startedAt, opacity })
      if (performance.now() - startedAt > 2200) return resolve({ passed: false, elapsed: performance.now() - startedAt, opacity, sawEndState, visibility: style.visibility })
      requestAnimationFrame(check)
    }
    requestAnimationFrame(check)
  })`)
  assert(reverseParent.passed, `A negative parent cTn@spd must play its child fade from the end toward the beginning: ${JSON.stringify(reverseParent)}`)

  const autoReverseParent = await evaluate(`new Promise(resolve => {
    const trigger = document.querySelector('.stage-shell > .slide-host .slide-element[title="Parent auto-reverse trigger"]')
    const target = document.querySelector('.stage-shell > .slide-host .slide-element[title="Parent auto-reverse fade"]')
    if (!trigger || !target) return resolve({ passed: false, reason: 'fixture shape missing' })
    let sawForwardEnd = false
    const startedAt = performance.now()
    trigger.click()
    const check = () => {
      const style = getComputedStyle(target)
      const opacity = Number(style.opacity)
      const visible = style.visibility !== 'hidden'
      if (visible && opacity > 0.85) sawForwardEnd = true
      if (sawForwardEnd && (!visible || opacity < 0.1)) return resolve({ passed: true, elapsed: performance.now() - startedAt, opacity })
      if (performance.now() - startedAt > 3200) return resolve({ passed: false, elapsed: performance.now() - startedAt, opacity, sawForwardEnd, visibility: style.visibility })
      requestAnimationFrame(check)
    }
    requestAnimationFrame(check)
  })`)
  assert(autoReverseParent.passed, `cTn@autoRev must play child progress forward and then return it to the initial state: ${JSON.stringify(autoReverseParent)}`)

  const childRepeatDuration = await evaluate(`new Promise(resolve => {
    const trigger = document.querySelector('.stage-shell > .slide-host .slide-element[title="Child repeat duration trigger"]')
    const target = document.querySelector('.stage-shell > .slide-host .slide-element[title="Child repeat duration fade"]')
    if (!trigger || !target) return resolve({ passed: false, reason: 'fixture shape missing' })
    let sawPeak = false
    let sawReset = false
    const startedAt = performance.now()
    trigger.click()
    const check = () => {
      const style = getComputedStyle(target)
      const opacity = Number(style.opacity)
      const visible = style.visibility !== 'hidden'
      if (visible && opacity > 0.85) sawPeak = true
      if (sawPeak && (!visible || opacity < 0.5)) sawReset = true
      if (sawReset && visible && opacity > 0.85) return resolve({ passed: true, elapsed: performance.now() - startedAt, opacity })
      if (performance.now() - startedAt > 1200) return resolve({ passed: false, elapsed: performance.now() - startedAt, opacity, sawPeak, sawReset, visibility: style.visibility })
      requestAnimationFrame(check)
    }
    requestAnimationFrame(check)
  })`)
  assert(childRepeatDuration.passed, `A finite child repeatDur without repeatCount must loop until its duration cap: ${JSON.stringify(childRepeatDuration)}`)

  const indefiniteRepeat = await evaluate(`new Promise(resolve => {
    const trigger = document.querySelector('.stage-shell > .slide-host .slide-element[title="Parent indefinite repeat trigger"]')
    const target = document.querySelector('.stage-shell > .slide-host .slide-element[title="Parent indefinite repeat fade"]')
    if (!trigger || !target) return resolve({ passed: false, reason: 'fixture shape missing' })
    let sawPeak = false
    let sawReset = false
    const startedAt = performance.now()
    trigger.click()
    const check = () => {
      const style = getComputedStyle(target)
      const opacity = Number(style.opacity)
      const visible = style.visibility !== 'hidden'
      if (visible && opacity > 0.85) sawPeak = true
      if (sawPeak && (!visible || opacity < 0.5)) sawReset = true
      if (sawReset && visible && opacity > 0.85) return resolve({ passed: true, elapsed: performance.now() - startedAt, opacity })
      if (performance.now() - startedAt > 2200) return resolve({ passed: false, elapsed: performance.now() - startedAt, opacity, sawPeak, sawReset, visibility: style.visibility })
      requestAnimationFrame(check)
    }
    requestAnimationFrame(check)
  })`)
  assert(indefiniteRepeat.passed, `A parent repeatDur="indefinite" must keep restarting its child animation: ${JSON.stringify(indefiniteRepeat)}`)

  const elementIterationStartedAt = await evaluate(`(() => { const started = performance.now(); document.querySelector('.stage-shell > .slide-host .slide-element[title="Element iteration first"]').click(); return started })()`)
  await delay(60)
  const firstElementIteration = await state()
  assert(firstElementIteration.slide?.startsWith('01 / 67') && Number(shape(firstElementIteration, 'Element iteration first').opacity) > 0 && Number(shape(firstElementIteration, 'Element iteration first').opacity) < 1, 'Element iteration must start the first grouped shape on its trigger click.')
  assert(shape(firstElementIteration, 'Element iteration second').visibility === 'hidden', 'Element iteration must retain its interval before revealing the next grouped shape.')
  await delay(190)
  const secondElementIteration = await state()
  assert(shape(secondElementIteration, 'Element iteration first').opacity === '1' && shape(secondElementIteration, 'Element iteration second').visibility === 'visible' && Number(shape(secondElementIteration, 'Element iteration second').opacity) > 0 && Number(shape(secondElementIteration, 'Element iteration second').opacity) < 1, `Element iteration must stagger child shapes by tmAbs: ${JSON.stringify({ elapsed: secondElementIteration.now - elementIterationStartedAt, first: shape(secondElementIteration, 'Element iteration first'), second: shape(secondElementIteration, 'Element iteration second'), warnings: secondElementIteration.warnings })}`)

  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(600)
  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(30)
  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(30)
  const secondSlide = await state()
  assert(secondSlide.slide?.startsWith('02 / 67'), 'The next click after all builds must advance the slide.')
  assert(secondSlide.slideTransition && secondSlide.slideTransition.keyframes[0] === 'translate(-100%, 100%)' && secondSlide.slideTransition.keyframes[1] === 'translate(0px, 0px)' && secondSlide.slideTransition.translateX < 0 && secondSlide.slideTransition.translateY > 0 && Number(secondSlide.slideTransition.duration) === 750, `The diagonal OOXML cover transition must follow its direction and duration: ${JSON.stringify(secondSlide.slideTransition)}`)
  assert(secondSlide.transitionUnderlay.slideCount === 1 && secondSlide.transitionUnderlay.shapeNames.includes('Click fade'), 'The previous slide must remain rendered below the incoming transition until it finishes.')
  assert(secondSlide.transitionUnderlay.translateX === 0 && secondSlide.transitionUnderlay.translateY === 0 && Math.abs(secondSlide.transitionUnderlay.frame.width - initial.hostFrame.width) < 1 && Math.abs(secondSlide.transitionUnderlay.frame.height - initial.hostFrame.height) < 1, `Cover must leave the outgoing page stationary beneath the incoming slide: ${JSON.stringify({ before: initial.hostFrame, underlay: secondSlide.transitionUnderlay })}`)
  await delay(850)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(850)
  assert(initial.warnings.some(warning => warning.includes('unsafe or unsupported URL'))
    && await evaluate(`document.querySelector('.stage-shell .slide-element[title="Unsafe text hyperlink"] .text-run-hyperlink') === null`),
  `Unsafe external hyperlink schemes must be ignored and reported: ${JSON.stringify(initial.warnings)}`)
  const internalLink = await evaluate(`(() => { const shape = document.querySelector('.stage-shell .slide-element[title="Go to slide 3"]'); return { role: shape?.getAttribute('role'), tabIndex: shape?.tabIndex } })()`)
  assert(internalLink.role === 'link' && internalLink.tabIndex === 0, `A linked shape must remain keyboard reachable: ${JSON.stringify(internalLink)}`)
  await evaluate(`(() => { const shape = document.querySelector('.stage-shell .slide-element[title="Go to slide 3"]'); shape?.focus(); shape?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })) })()`)
  await delay(100)
  const keyboardInternalJump = await state()
  assert(keyboardInternalJump.slide?.startsWith('03 / 67'), `Enter on a focused linked shape must navigate to its target slide: ${keyboardInternalJump.slide}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(850)
  await evaluate(`document.querySelector('.stage-shell .slide-element[title="Go to slide 3"]')?.click()`)
  await delay(100)
  const internalJump = await state()
  assert(internalJump.slide?.startsWith('03 / 67'), `A shape hlinkClick relationship must navigate directly to its slide and stop click-to-advance: ${internalJump.slide}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(850)
  const openedExternalLink = await evaluate(`(() => { const open = window.open; let result; window.open = (...args) => { result = args; return null }; document.querySelector('.stage-shell .text-run-hyperlink[title="Open external documentation"]')?.click(); window.open = open; return result })()`)
  await delay(80)
  const afterExternalLink = await state()
  assert(openedExternalLink?.[0] === 'https://example.test/reference' && openedExternalLink?.[1] === '_blank' && openedExternalLink?.[2] === 'noopener,noreferrer'
    && afterExternalLink.slide?.startsWith('01 / 67'), `A text hyperlink must open its safe external target without advancing the slide: ${JSON.stringify({ openedExternalLink, slide: afterExternalLink.slide })}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(850)
  const chartRender = await evaluate(`(() => { const format = (value, options) => new Intl.NumberFormat(undefined, options).format(value); const percent = value => format(value, { style: 'percent', maximumFractionDigits: 1 }); return [...document.querySelectorAll('.slide-chart')].map(svg => { const name = svg.closest('.slide-element')?.title; return { name, title: svg.querySelector('title')?.textContent, labels: [...svg.querySelectorAll('text')].map(item => item.textContent), dataLabels: [...svg.querySelectorAll('.chart-data-label')].map(item => [...item.querySelectorAll('tspan')].map(line => line.textContent).join('\\n')), expectedDataLabels: name === 'Quarterly sales chart' ? [format(9, { useGrouping: true, minimumFractionDigits: 1, maximumFractionDigits: 1 }), format(18, { useGrouping: false, minimumFractionDigits: 2, maximumFractionDigits: 2 }), format(15, { useGrouping: true, minimumFractionDigits: 1, maximumFractionDigits: 1 }), format(21, { useGrouping: true, minimumFractionDigits: 1, maximumFractionDigits: 1 })] : name === 'Quarterly trend chart' ? [18, 15, 26, 21].map(value => format(value, { style: 'percent', useGrouping: false, minimumFractionDigits: 1, maximumFractionDigits: 1 })) : name === 'Quarterly share chart' ? ['Q1\\n' + percent(12 / 56), '18 / Q2', 'Q3\\n' + percent(26 / 56)] : undefined, bars: [...svg.querySelectorAll('rect')].filter(item => Number(item.getAttribute('height')) > 20).map(item => item.getAttribute('fill')), paths: svg.querySelectorAll('path').length, markers: svg.querySelectorAll('circle').length }; }) })()`)
  const barChart = chartRender.find(chart => chart.name === 'Quarterly sales chart')
  const lineChart = chartRender.find(chart => chart.name === 'Quarterly trend chart')
  const pieChart = chartRender.find(chart => chart.name === 'Quarterly share chart')
  const doughnutGeometry = await evaluate(`(() => { const svg = [...document.querySelectorAll('.slide-chart')].find(item => item.closest('.slide-element')?.title === 'Quarterly allocation doughnut'); const path = svg?.querySelector('path'); const d = path?.getAttribute('d') || ''; return { count: svg?.querySelectorAll('path').length, fillRule: path?.getAttribute('fill-rule'), start: /^M\\s+([-.\\d]+)\\s+([-.\\d]+)/.exec(d)?.slice(1).map(Number), arcRadii: [...d.matchAll(/A\\s+([-.\\d]+)\\s+([-.\\d]+)\\s+0\\s+[01]\\s+[01]/g)].map(match => Number(match[1])) }; })()`)
  const valueAxisRender = await evaluate(`(() => [...document.querySelectorAll('.slide-chart')].map(svg => ({ name: svg.closest('.slide-element')?.title, ticks: [...svg.querySelectorAll('.chart-value-axis-label')].map(item => ({ label: item.textContent, x: Number(item.getAttribute('x')), y: Number(item.getAttribute('y')), anchor: item.getAttribute('text-anchor') })) })))()`)
  const barValueAxis = valueAxisRender.find(axis => axis.name === 'Quarterly sales chart')
  const lineValueAxis = valueAxisRender.find(axis => axis.name === 'Quarterly trend chart')
  const chartAxisCrossings = await evaluate(`(() => [...document.querySelectorAll('.slide-chart')].map(svg => { const line = selector => { const item = svg.querySelector(selector); return item && { x1: Number(item.getAttribute('x1')), y1: Number(item.getAttribute('y1')), x2: Number(item.getAttribute('x2')), y2: Number(item.getAttribute('y2')) } }; return { name: svg.closest('.slide-element')?.title, category: line('.chart-category-axis-line'), value: line('.chart-value-axis-line'), gridlines: svg.querySelectorAll('.chart-major-gridline').length } }))()`)
  const chartAxisTicks = await evaluate(`(() => [...document.querySelectorAll('.slide-chart')].map(svg => ({ name: svg.closest('.slide-element')?.title, ticks: [...svg.querySelectorAll('.chart-axis-tick')].map(item => ({ axis: item.getAttribute('data-axis'), x1: Number(item.getAttribute('x1')), y1: Number(item.getAttribute('y1')), x2: Number(item.getAttribute('x2')), y2: Number(item.getAttribute('y2')) })) })))()`)
  const horizontalChartGeometry = await evaluate(`(() => { const svg = [...document.querySelectorAll('.slide-chart')].find(item => item.closest('.slide-element')?.title === 'Horizontal axis tick chart'); return { bars: [...(svg?.querySelectorAll('rect') || [])].filter(item => Number(item.getAttribute('height')) > 20).map(item => ({ width: Number(item.getAttribute('width')), height: Number(item.getAttribute('height')) })) } })()`)
  const categoryAxisRender = await evaluate(`(() => [...document.querySelectorAll('.slide-chart')].map(svg => ({ name: svg.closest('.slide-element')?.title, labels: [...svg.querySelectorAll('.chart-category-label')].map(item => ({ text: item.textContent, x: Number(item.getAttribute('x')), y: Number(item.getAttribute('y')), anchor: item.getAttribute('text-anchor') })) })))()`)
  const lineCategoryAxis = categoryAxisRender.find(axis => axis.name === 'Quarterly trend chart')
  const barCategoryAxis = categoryAxisRender.find(axis => axis.name === 'Quarterly sales chart')
  const horizontalCategoryAxis = categoryAxisRender.find(axis => axis.name === 'Horizontal axis tick chart')
  const chartLegends = await evaluate(`(() => [...document.querySelectorAll('.slide-chart')].map(svg => ({ name: svg.closest('.slide-element')?.title, items: [...svg.querySelectorAll('.chart-legend-label')].map(item => ({ text: item.textContent, x: Number(item.getAttribute('x')), y: Number(item.getAttribute('y')) })) })))()`)
  const manualLegendClip = await evaluate(`(() => { const svg = [...document.querySelectorAll('.slide-chart')].find(item => item.closest('.slide-element')?.title === 'Value axis deleted chart'); const group = svg?.querySelector('.chart-legend-group'); const id = /url\\(#([^)]+)\\)/.exec(group?.getAttribute('clip-path') || '')?.[1]; const rect = [...(svg?.querySelectorAll('clipPath') || [])].find(item => item.id === id)?.querySelector('rect'); return rect && { x: Number(rect.getAttribute('x')), y: Number(rect.getAttribute('y')), width: Number(rect.getAttribute('width')), height: Number(rect.getAttribute('height')) } })()`)
  const dataLabelGeometry = await evaluate(`(() => [...document.querySelectorAll('.slide-chart')].map(svg => ({ name: svg.closest('.slide-element')?.title, labels: [...svg.querySelectorAll('.chart-data-label')].map(item => ({ text: item.textContent, x: Number(item.getAttribute('x')), y: Number(item.getAttribute('y')), anchor: item.getAttribute('text-anchor') })) })))()`)
  const barLegend = chartLegends.find(legend => legend.name === 'Quarterly sales chart')
  const lineLegend = chartLegends.find(legend => legend.name === 'Quarterly trend chart')
  const pieLegend = chartLegends.find(legend => legend.name === 'Quarterly share chart')
  const barPositionedLabels = dataLabelGeometry.find(chart => chart.name === 'Quarterly sales chart')
  const linePositionedLabels = dataLabelGeometry.find(chart => chart.name === 'Quarterly trend chart')
  assert(barChart?.title === 'Quarterly sales' && ['Q1', 'Q2', 'Q3', 'South'].every(label => barChart.labels.includes(label)) && !barChart.labels.includes('North') && barChart.bars.length === 6 && new Set(barChart.bars).size === 2, `A linked chart must render cached categories, preserve both series, and apply legend-entry deletion: ${JSON.stringify(chartRender)}`)
  assert(barValueAxis?.ticks.map(tick => tick.label).join(',') === '0.0,10.0,20.0,30.0', `Value-axis min, max, major unit, and number format must render: ${JSON.stringify(valueAxisRender)}`)
  assert(barValueAxis?.ticks.every(tick => tick.x === 78 && tick.anchor === 'end') && barCategoryAxis?.labels.length === 3 && barCategoryAxis.labels.every(label => label.y === 65 && label.anchor === 'middle'), `c:tickLblPos high/low must place category and value labels at the perpendicular axis ends: ${JSON.stringify({ value: barValueAxis, category: barCategoryAxis })}`)
  assert(barLegend?.items.map(item => item.text).join(',') === 'South' && barLegend.items.every(item => item.y === 530), `c:legendEntry/c:delete must hide the indexed series entry without moving the remaining bottom legend: ${JSON.stringify(chartLegends)}`)
  assert(lineLegend?.items[0]?.x === 118 && lineLegend.items[0]?.y === 135 && lineLegend.items[1]?.x === 118 && lineLegend.items[1]?.y === 163, `Chart legend manual edge/factor coordinates must set the origin and preserve vertical spacing: ${JSON.stringify(chartLegends)}`)
  assert(Math.abs(barPositionedLabels?.labels.find(label => label.text === '18.00')?.y - 354) < 1, `A point-level c:dLblPos must override the series position and center its column label: ${JSON.stringify(dataLabelGeometry)}`)
  assert(barChart.dataLabels.join(',') === barChart.expectedDataLabels.join(','), `Chart, series, and point overrides must resolve while applying decimal/grouping formats: ${JSON.stringify(barChart)}`)
  assert(lineChart?.markers === 6 && lineChart.paths === 6, `Line chart caches must render series paths and point markers: ${JSON.stringify(lineChart)}`)
  assert(lineValueAxis?.ticks.map(tick => tick.label).join(',') === '10.0,15.0,20.0,25.0,30.0' && lineValueAxis.ticks[0].y < lineValueAxis.ticks.at(-1).y, `A reversed value axis must place minimum above maximum: ${JSON.stringify(valueAxisRender)}`)
  assert(lineLegend?.items.map(item => item.text).join(',') === 'North,South' && lineLegend.items.every(item => item.x === 118 && item.y >= 135) && lineValueAxis.ticks.every(tick => tick.x === 644 && tick.anchor === 'start'), `A c:catAx/crossesAt category index must position the value-axis labels at that category: ${JSON.stringify({ chartLegends, valueAxisRender })}`)
  assert(lineCategoryAxis?.labels.map(label => label.text).join(',') === 'Q1,Q2,Q3' && lineCategoryAxis.labels.every(label => Math.abs(label.y - 162.5) < 0.01 && label.anchor === 'middle'), `A c:valAx/crossesAt value must position the category-axis labels at that value: ${JSON.stringify(categoryAxisRender)}`)
  assert(valueAxisRender.find(axis => axis.name === 'Horizontal axis tick chart')?.ticks.length === 0 && horizontalCategoryAxis?.labels.length === 3 && horizontalCategoryAxis.labels.every(label => label.x === 78 && label.anchor === 'end'), `c:tickLblPos none must hide value labels while low positions horizontal-bar category labels at the low value end: ${JSON.stringify({ value: valueAxisRender, category: horizontalCategoryAxis })}`)
  const barAxisCrossings = chartAxisCrossings.find(chart => chart.name === 'Quarterly sales chart')
  const lineAxisCrossings = chartAxisCrossings.find(chart => chart.name === 'Quarterly trend chart')
  const barAxisTicks = chartAxisTicks.find(chart => chart.name === 'Quarterly sales chart')?.ticks || []
  const lineAxisTicks = chartAxisTicks.find(chart => chart.name === 'Quarterly trend chart')?.ticks || []
  const horizontalAxisTicks = chartAxisTicks.find(chart => chart.name === 'Horizontal axis tick chart')?.ticks || []
  const categoryDeletedAxis = chartAxisCrossings.find(chart => chart.name === 'Category axis deleted chart')
  const valueDeletedAxis = chartAxisCrossings.find(chart => chart.name === 'Value axis deleted chart')
  const horizontalAxisCrossings = chartAxisCrossings.find(chart => chart.name === 'Horizontal axis tick chart')
  const categoryDeletedTicks = chartAxisTicks.find(chart => chart.name === 'Category axis deleted chart')?.ticks || []
  const valueDeletedTicks = chartAxisTicks.find(chart => chart.name === 'Value axis deleted chart')?.ticks || []
  const valueDeletedOverlayLegend = chartLegends.find(legend => legend.name === 'Value axis deleted chart')
  const categoryDeletedLabels = categoryAxisRender.find(axis => axis.name === 'Category axis deleted chart')?.labels || []
  const valueDeletedLabels = valueAxisRender.find(axis => axis.name === 'Value axis deleted chart')?.ticks || []
  assert(barAxisCrossings?.category?.y1 === 75 && barAxisCrossings.category.y2 === 75 && barAxisCrossings.value?.x1 === 938, `c:crosses=max must place the category and value axis lines at their maximum crossings: ${JSON.stringify(chartAxisCrossings)}`)
  assert(Math.abs((lineAxisCrossings?.category?.y1 ?? NaN) - 172.5) < 0.01 && lineAxisCrossings?.value?.x1 === 634, `Both c:crossesAt values must position the intersecting axis lines: ${JSON.stringify(chartAxisCrossings)}`)
  assert(barAxisTicks.filter(tick => tick.axis === 'category').length === 0 && barAxisTicks.filter(tick => tick.axis === 'value').length === 4 && barAxisTicks.filter(tick => tick.axis === 'value').every(tick => tick.y1 === tick.y2 && tick.x1 === 938 && tick.x2 === 933), `c:majorTickMark val="out" must draw outward value-axis marks while "none" suppresses category marks: ${JSON.stringify(barAxisTicks)}`)
  assert(lineAxisTicks.filter(tick => tick.axis === 'category').length === 3 && lineAxisTicks.filter(tick => tick.axis === 'category').every(tick => tick.x1 === tick.x2 && Math.abs(tick.y1 - 167.5) < 0.01 && Math.abs(tick.y2 - 177.5) < 0.01) && lineAxisTicks.filter(tick => tick.axis === 'value').length === 5 && lineAxisTicks.filter(tick => tick.axis === 'value').every(tick => tick.x1 === 634 && tick.x2 === 629 && tick.y1 === tick.y2), `c:majorTickMark val="cross" and "in" must follow their axis sides and crossings: ${JSON.stringify(lineAxisTicks)}`)
  assert(horizontalChartGeometry.bars.length === 6 && horizontalChartGeometry.bars.every(bar => bar.width > bar.height) && horizontalAxisTicks.filter(tick => tick.axis === 'category').length === 3 && horizontalAxisTicks.filter(tick => tick.axis === 'category').every(tick => tick.y1 === tick.y2 && tick.x1 === 938 && tick.x2 === 943) && horizontalAxisTicks.filter(tick => tick.axis === 'value').length === 4 && horizontalAxisTicks.filter(tick => tick.axis === 'value').every(tick => tick.x1 === tick.x2 && tick.y1 === 75 && tick.y2 === 80), `Horizontal bar charts must render horizontal bars and category/value ticks at right/top axes: ${JSON.stringify({ horizontalChartGeometry, horizontalAxisTicks })}`)
  assert(!categoryDeletedAxis?.category && !!categoryDeletedAxis?.value && categoryDeletedAxis.gridlines === 4 && categoryDeletedTicks.filter(tick => tick.axis === 'category').length === 0 && categoryDeletedTicks.filter(tick => tick.axis === 'value').length === 4 && categoryDeletedLabels.length === 0, `c:catAx/delete must remove only the category axis while retaining value-axis gridlines: ${JSON.stringify({ axis: categoryDeletedAxis, ticks: categoryDeletedTicks, labels: categoryDeletedLabels })}`)
  assert(!!valueDeletedAxis?.category && !valueDeletedAxis?.value && valueDeletedAxis.gridlines === 0 && valueDeletedTicks.filter(tick => tick.axis === 'category').length === 3 && valueDeletedTicks.filter(tick => tick.axis === 'value').length === 0 && valueDeletedLabels.length === 0, `c:valAx/delete must remove the value axis, its labels, ticks, and gridlines while retaining category visuals: ${JSON.stringify({ axis: valueDeletedAxis, ticks: valueDeletedTicks, labels: valueDeletedLabels })}`)
  assert(valueDeletedOverlayLegend?.items.map(item => item.text).join(',') === 'North,South' && valueDeletedAxis?.category?.x1 === 938 && valueDeletedAxis.category.x2 === 938, `c:legend/c:overlay must draw the legend without reserving plot-area space: ${JSON.stringify({ legend: valueDeletedOverlayLegend, axis: valueDeletedAxis })}`)
  assert(manualLegendClip?.x === 720 && manualLegendClip.y === 105 && manualLegendClip.width === 230 && manualLegendClip.height === 300, `c:manualLayout w/h edge and factor modes must produce the expected legend clip box: ${JSON.stringify(manualLegendClip)}`)
  assert(horizontalAxisCrossings?.gridlines === 0 && !!horizontalAxisCrossings.category && !!horizontalAxisCrossings.value, `An absent c:majorGridlines must hide gridlines while preserving both axes: ${JSON.stringify(horizontalAxisCrossings)}`)
  const rightPointLabel = linePositionedLabels?.labels.find(label => label.text === '2600.0% | North')
  assert(rightPointLabel?.anchor === 'start' && rightPointLabel.x > 840, `A series-level c:dLblPos must override the chart label position: ${JSON.stringify(dataLabelGeometry)}`)
  assert(lineChart.dataLabels.filter(label => label.includes(' | ')).map(label => label.split(' | ')[0]).join(',') === lineChart.expectedDataLabels.join(',') && ['North', 'South'].every(label => lineChart.dataLabels.includes(label)) && lineChart.dataLabels.filter(label => label.includes(' | ')).every(label => / \| (North|South)$/.test(label)), `Series names and custom separators must follow chart and point overrides: ${JSON.stringify(lineChart)}`)
  assert(pieChart?.paths === 3 && ['Q1', 'Q2', 'Q3'].every(label => pieChart.dataLabels.some(item => item.includes(label))), `Pie chart caches must render one colored slice and data label per category: ${JSON.stringify(pieChart)}`)
  assert(doughnutGeometry.count === 3 && doughnutGeometry.fillRule === 'evenodd' && doughnutGeometry.arcRadii.length === 2 && Math.abs(doughnutGeometry.arcRadii[1] / doughnutGeometry.arcRadii[0] - 0.4) < 0.001 && doughnutGeometry.start?.[1] === 270 && doughnutGeometry.start?.[0] > 513, `Doughnut charts must render annular slices with configured hole size and first-slice angle: ${JSON.stringify(doughnutGeometry)}`)
  assert(pieLegend?.items.map(item => item.text).join(',') === 'Q1,Q3', `c:legendEntry/c:delete must hide the indexed pie-point entry: ${JSON.stringify(chartLegends)}`)
  assert(chartLegends.find(legend => legend.name === 'Horizontal axis tick chart')?.items.length === 0, `Charts without c:legend must not synthesize a legend: ${JSON.stringify(chartLegends)}`)
  assert(pieChart.dataLabels.join(',') === pieChart.expectedDataLabels.join(','), `Pie percentage labels and point-specific value overrides must resolve: ${JSON.stringify(pieChart)}`)
  assert(secondSlide.tables.length === 1 && secondSlide.tables[0].cells.length === 5, 'The basic table, merged-cell layout, and text-column row must render.')
  assert(secondSlide.tables[0].cells[0].text === 'Merged header' && secondSlide.tables[0].cells[0].colSpan === 2, 'The table header must span both columns.')
  assert(secondSlide.tables[0].cells[0].backgroundColor === 'rgb(30, 64, 175)' && secondSlide.tables[0].cells[0].borderTopWidth !== '0px', 'Direct table fills and borders must render.')
  const dynamicTableCell = secondSlide.tables[0].cells.find(cell => cell.autoFit === 'normal')
  const referenceTableCell = secondSlide.tables[0].cells.find(cell => cell.text === 'Left cell')
  assert(dynamicTableCell && referenceTableCell && dynamicTableCell.fontSize < referenceTableCell.fontSize
    && dynamicTableCell.scrollHeight <= dynamicTableCell.clientHeight,
    `A table-cell a:normAutofit must shrink text to fit its saved row height: ${JSON.stringify({ dynamicTableCell, referenceTableCell })}`)
  const columnTableCellRender = secondSlide.tables[0].cells.find(cell => cell.text.includes('fills its first text column'))
  const expectedTableColumnGap = 457200 * (columnTableCellRender?.frameWidth || 0) / 12192000
  const tableColumnXs = (columnTableCellRender?.runRects || []).map(rect => rect.left)
  const tableColumnFlow = await evaluate(`(() => {
    const cell = [...document.querySelectorAll('.stage-shell > .slide-host .slide-table td')].find(item => item.textContent.includes('fills its first text column'))
    const node = cell?.querySelector('.text-run')?.firstChild
    if (!node || node.nodeType !== Node.TEXT_NODE) return undefined
    const rectAt = (start, end) => { const range = document.createRange(); range.setStart(node, start); range.setEnd(node, end); const rect = range.getBoundingClientRect(); return { left: rect.left, right: rect.right } }
    const rect = element => { const { left, right, top, width, height } = element.getBoundingClientRect(); return { left, right, top, width, height } }
    return { first: rectAt(0, 5), last: rectAt(Math.max(0, node.textContent.length - 5), node.textContent.length), cell: rect(cell), frame: rect(cell.querySelector('.table-cell-text')), paragraphRects: [...cell.querySelector('.text-paragraph').getClientRects()].map(item => ({ left: item.left, top: item.top, width: item.width, height: item.height })) }
  })()`)
  assert(columnTableCellRender?.columnCount === '2' && columnTableCellRender.direction === 'rtl' && columnTableCellRender.paragraphDirection === 'ltr' && Math.abs(Number.parseFloat(columnTableCellRender.columnGap) - expectedTableColumnGap) < 0.1 && tableColumnXs.length > 1 && Math.max(...tableColumnXs) - Math.min(...tableColumnXs) > expectedTableColumnGap && tableColumnFlow?.first.left > tableColumnFlow?.last.left, `Table-cell a:bodyPr rtlCol must reverse column order while preserving text direction and EMU spacing: ${JSON.stringify({ columnTableCellRender: { ...columnTableCellRender, runRects: columnTableCellRender.runRects?.slice(0, 4) }, tableColumnFlow, expectedTableColumnGap })}`)
  const customRadialMetrics = await evaluate(`(() => {
    const path = document.querySelector('.stage-shell > .slide-host .slide-element[title="Custom radial gradient geometry"] .custom-geometry path')
    const read = property => {
      const paint = path?.getAttribute(property) || ''
      const id = /^url\\(#(.+)\\)$/.exec(paint)?.[1]
      const gradient = id ? document.getElementById(id) : undefined
      return { paint, type: gradient?.localName, units: gradient?.getAttribute('gradientUnits'), cx: Number(gradient?.getAttribute('cx')), cy: Number(gradient?.getAttribute('cy')), r: Number(gradient?.getAttribute('r')), stops: [...(gradient?.querySelectorAll('stop') || [])].map(stop => ({ offset: stop.getAttribute('offset'), color: stop.getAttribute('stop-color'), opacity: stop.getAttribute('stop-opacity') })) }
    }
    return { fill: read('fill'), stroke: read('stroke') }
  })()`)
  assert(customRadialMetrics.fill.type === 'radialGradient' && customRadialMetrics.fill.units === 'userSpaceOnUse'
    && customRadialMetrics.fill.cx === 500 && customRadialMetrics.fill.cy === 500 && customRadialMetrics.fill.r === 500
    && customRadialMetrics.fill.stops.length === 3 && customRadialMetrics.fill.stops[0]?.opacity === '0.5'
    && customRadialMetrics.stroke.type === 'radialGradient' && customRadialMetrics.stroke.stops.length === 2,
  `Custom-geometry path gradients must preserve all radial fill and outline stops, including alpha: ${JSON.stringify(customRadialMetrics)}`)
  assert(shape(secondSlide, 'Gradient shape').backgroundImage.includes('linear-gradient') && shape(secondSlide, 'Gradient shape').backgroundImage.includes('rgba(37, 99, 235, 0.5)'), 'Gradient stops, angle, and alpha must render as a CSS gradient.')
  assert(shape(secondSlide, 'Gradient shape').boxShadow !== 'none', 'Shape outer shadows must render.')
  assert(shape(secondSlide, 'Luminance transform').backgroundColor === 'rgb(0, 230, 0)', 'lumMod/lumOff must modify HSL luminance while preserving the saturated green hue.')
  const grouped = shape(secondSlide, 'Rotated group child')
  assert(grouped, 'Grouped shape children must render.')
  assert(Math.abs(grouped.frameRect.width / grouped.frameRect.frameWidth - 2_000_000 / 12_192_000) < 0.005, 'The group rotation must preserve the child bounds width.')
  assert(Math.abs(grouped.frameRect.height / grouped.frameRect.frameHeight - 4_000_000 / 6_858_000) < 0.005, 'The group rotation and scale must transform child bounds.')
  assert(Math.abs(grouped.frameRect.left / grouped.frameRect.frameWidth - 6_500_000 / 12_192_000) < 0.005, 'Group child offsets and the rotated group center must map to slide coordinates.')
  const nested = shape(secondSlide, 'Nested group child')
  assert(nested, 'Nested group children must render.')
  assert(Math.abs(nested.frameRect.width / nested.frameRect.frameWidth - 800_000 / 12_192_000) < 0.005, 'Nested group scaling and parent rotation must compose.')
  assert(Math.abs(nested.frameRect.left / nested.frameRect.frameWidth - 7_500_000 / 12_192_000) < 0.005, 'Nested group offsets and horizontal flips must preserve the expected bounds.')
  assert(!secondSlide.warnings.some(warning => warning.includes('Grouped shapes')), 'Supported groups must not be reported as omitted content.')
  await delay(800)
  assert((await state()).transitionUnderlay.slideCount === 0, 'The outgoing slide snapshot must be removed when the transition finishes.')
  if (process.env.DECKLINE_CAPTURE_SLIDE_PATH) {
    const rect = await evaluate(`(() => { const { x, y, width, height } = document.querySelector('.stage-shell > .slide-host .slide-frame').getBoundingClientRect(); return { x, y, width, height } })()`)
    const screenshot = await call('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true, clip: { ...rect, scale: 1280 / rect.width } })
    await writeFile(process.env.DECKLINE_CAPTURE_SLIDE_PATH, Buffer.from(screenshot.data, 'base64'))
  }

  await evaluate(`document.querySelector('.player-controls button:first-child').click()`)
  await delay(50)
  const returned = await state()
  assert(returned.slide?.startsWith('01 / 67'), 'Previous must return to the first slide.')
  assert(returned.slideTransition && returned.slideTransition.keyframes[0] === 'scale(0.5)' && returned.slideTransition.keyframes[1] === 'scale(1)', 'Returning to the first slide must play its OOXML zoom transition.')
  assert(shape(returned, 'Click fade').visibility === 'visible' && shape(returned, 'With previous').visibility === 'visible' && shape(returned, 'Second click').visibility === 'visible', 'Returning must show completed entrance builds.')
  assert(paragraph(returned, 'Built paragraph').visibility === 'visible', 'Returning must show completed paragraph builds.')
  assert(shape(returned, 'Exit after previous').visibility === 'visible', 'Returning must preserve the later sibling animation that resets the frozen exit fill.')

  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(50)
  const numberJump = await state()
  assert(numberJump.slide?.startsWith('02 / 67'), 'Number plus Enter must jump to that slide.')
  await evaluate(`document.querySelector('.stage-shell').dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 120 }))`)
  await delay(20)
  const menuOpen = await state()
  assert(menuOpen.menuItems.join('|') === 'Previous|Next|Black screen|White screen', 'Right-click must open the presentation menu.')
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'b' }))`)
  await delay(20)
  const blackScreen = await state()
  assert(blackScreen.menuItems.length === 0 && blackScreen.screenOverlay?.includes('is-black'), 'Escape must close the menu and B must show a black screen.')
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(50)
  const pushSlide = await state()
  assert(pushSlide.slide?.startsWith('03 / 67') && pushSlide.slideTransition?.keyframes[0] === 'translateX(-100%)' && pushSlide.slideTransition.translateX < 0, `The push transition must move the incoming slide left-to-right from the OOXML direction: ${JSON.stringify({ slide: pushSlide.slide, transition: pushSlide.slideTransition, underlay: pushSlide.transitionUnderlay })}`)
  assert(blackScreen.slideTransition && pushSlide.transitionUnderlay.translateX > blackScreen.slideTransition.translateX && pushSlide.transitionUnderlay.translateX < blackScreen.slideTransition.translateX + initial.hostFrame.width, `The paired push must move the outgoing snapshot right from the in-progress cover position: ${JSON.stringify({ cover: blackScreen.slideTransition, outgoing: pushSlide.transitionUnderlay })}`)
  await delay(800)
  const beforePull = await state()
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(70)
  const pullSlide = await state()
  assert(pullSlide.slide?.startsWith('04 / 67') && !pullSlide.slideTransition, 'Pull must animate the retained outgoing slide while the incoming slide remains stationary.')
  assert(pullSlide.transitionUnderlay.pulling && Number(pullSlide.transitionUnderlay.layerZIndex) > Number(pullSlide.transitionUnderlay.hostZIndex), 'The pull snapshot must be stacked above the incoming slide.')
  assert(pullSlide.transitionUnderlay.shapeNames.includes('Third slide') && pullSlide.transitionUnderlay.keyframes?.at(-1)?.includes('translate(-100%, 100%)') && Number(pullSlide.transitionUnderlay.duration) === 750, `The pull snapshot must move left and down for dir=ld: ${JSON.stringify(pullSlide.transitionUnderlay)}`)
  assert(pullSlide.transitionUnderlay.frame.left < beforePull.hostFrame.left - 5 && pullSlide.transitionUnderlay.frame.top > beforePull.hostFrame.top + 5, `The old page must visibly move left and down while revealing the new slide: ${JSON.stringify({ before: beforePull.hostFrame, outgoing: pullSlide.transitionUnderlay.frame })}`)
  await delay(800)
  assert((await state()).transitionUnderlay.slideCount === 0, 'The pulled-off outgoing snapshot must be removed when the transition finishes.')
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '5' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const splitOut = await state()
  assert(splitOut.slide?.startsWith('05 / 67') && splitOut.slideTransition?.clipKeyframes?.[0]?.includes('50%') && splitOut.slideTransition.clipKeyframes.at(-1)?.includes('0'), `Vertical split-out must reveal the new page from its center: ${JSON.stringify(splitOut.slideTransition)}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '6' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const splitIn = await state()
  assert(splitIn.slide?.startsWith('06 / 67') && !splitIn.slideTransition && splitIn.transitionUnderlay.splittingIn && Number(splitIn.transitionUnderlay.layerZIndex) > Number(splitIn.transitionUnderlay.hostZIndex), 'Horizontal split-in must animate the old page above the new page.')
  assert(splitIn.transitionUnderlay.shapeNames.includes('Split out slide') && splitIn.transitionUnderlay.clipKeyframes?.at(-1)?.includes('50% 0') && Number(splitIn.transitionUnderlay.duration) === 750, `Horizontal split-in must close the outgoing page toward its center: ${JSON.stringify(splitIn.transitionUnderlay)}`)
  await delay(800)
  assert((await state()).transitionUnderlay.slideCount === 0, 'Split must remove the outgoing snapshot after the transition.')
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '7' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const circle = await state()
  assert(circle.slide?.startsWith('07 / 67') && circle.slideTransition?.clipKeyframes?.[0]?.startsWith('circle(0%') && circle.slideTransition.clipKeyframes.at(-1)?.startsWith('circle(100%'), `Circle must grow from the slide center to reveal the new page: ${JSON.stringify(circle.slideTransition)}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '8' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const diamond = await state()
  assert(diamond.slide?.startsWith('08 / 67') && diamond.slideTransition?.clipKeyframes?.[0]?.includes('50% 50%') && diamond.slideTransition.clipKeyframes.at(-1)?.includes('-50%'), `Diamond must expand beyond slide edges: ${JSON.stringify(diamond.slideTransition)}`)
  await delay(800)
  assert((await state()).transitionUnderlay.slideCount === 0, 'The shape transitions must leave no outgoing snapshot behind.')
  await evaluate(`window.__decklineCutBlackObserved = false; window.__decklineCutBlackObserver = new MutationObserver(() => { if (document.querySelector('.stage-shell')?.classList.contains('is-cut-through-black')) window.__decklineCutBlackObserved = true }); window.__decklineCutBlackObserver.observe(document.querySelector('.stage-shell'), { attributes: true, attributeFilter: ['class'] })`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '9' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(30)
  const cutThroughBlack = await state()
  assert(cutThroughBlack.slide?.startsWith('09 / 67') && !cutThroughBlack.slideTransition && cutThroughBlack.cutBlackObserved && !cutThroughBlack.warnings.some(warning => warning.includes('through-black')) && cutThroughBlack.warnings.some(warning => warning.includes('AlternateContent')), 'An unsupported AlternateContent choice must render its through-black cut fallback and report only the unsupported-choice warning.')
  await evaluate(`window.__decklineCutBlackObserver.disconnect(); delete window.__decklineCutBlackObserver`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '0' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(30)
  const cut = await state()
  assert(cut.slide?.startsWith('10 / 67') && !cut.slideTransition && cut.transitionUnderlay.slideCount === 0, 'Cut must switch slides immediately without an animation or retained overlay.')
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const honeycomb = await state()
  const honeycombStart = honeycomb.slideTransition?.clipKeyframes?.[0] || ''
  const honeycombEnd = honeycomb.slideTransition?.clipKeyframes?.at(-1) || ''
  assert(honeycomb.slide?.startsWith('11 / 67') && Number(honeycomb.slideTransition?.duration) === 900 && honeycombStart.startsWith('path("M') && honeycombEnd.startsWith('path("M') && (honeycombEnd.match(/M(?=\s*[-\d])/g) || []).length > 50 && honeycomb.slideTransition.clipPath.startsWith('path(') && !honeycomb.warnings.some(warning => warning.includes('honeycomb')), 'The Office 2010 honeycomb transition must render as an animated tiled hexagon mask without a fallback warning and honor p14:dur.')
  await delay(950)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const blinds = await state()
  assert(blinds.slide?.startsWith('12 / 67') && blinds.slideTransition?.clipKeyframes?.length === 3 && blinds.slideTransition.clipKeyframes[0]?.startsWith('path("M') && blinds.slideTransition.clipKeyframes.at(-1)?.startsWith('path("M') && blinds.slideTransition.clipPath.startsWith('path(') && !blinds.warnings.some(warning => warning.includes('blinds')), 'Vertical blinds must reveal the new slide through animated, full-height slats without a fallback warning.')
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const checker = await state()
  const checkerEnd = checker.slideTransition?.clipKeyframes?.at(-1) || ''
  assert(checker.slide?.startsWith('13 / 67') && checker.slideTransition?.clipKeyframes?.length === 5 && checker.slideTransition.clipKeyframes[0]?.startsWith('path("M') && (checkerEnd.match(/M(?=\s*[-\d])/g) || []).length > 50 && checker.slideTransition.clipPath.startsWith('path(') && !checker.warnings.some(warning => warning.includes('checker')), 'Horizontal checkerboard must reveal through a tiled square mask without a fallback warning.')
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const wheel = await state()
  const wheelEnd = wheel.slideTransition?.clipKeyframes?.at(-1) || ''
  const reverseBoundary = wheel.slideTransition?.clipKeyframes?.[1]?.split(' L')[1]?.trim().split(' ').slice(0, 2).map(Number)
  const reverseStartAngle = reverseBoundary && Math.atan2(reverseBoundary[1] - wheel.hostFrame.height / 2, reverseBoundary[0] - wheel.hostFrame.width / 2)
  assert(wheel.slide?.startsWith('14 / 67') && wheel.slideTransition?.clipKeyframes?.length === 5 && (wheelEnd.match(/M(?=\s*[-\d])/g) || []).length === 8 && wheel.slideTransition.clipPath.startsWith('path(') && Number(wheel.slideTransition.duration) === 750, 'The eight-spoke wheel transition must render eight animated radial wedges.')
  assert(reverseStartAngle !== undefined && Math.abs(reverseStartAngle + Math.PI / 4) < 0.05, `The p14 eight-spoke reverse wheel must select its AlternateContent choice and sweep from the opposite spoke boundary: ${JSON.stringify({ reverseStartAngle, reverseBoundary, warnings: wheel.warnings })}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(800)
  await evaluate(`document.querySelector('.stage-shell').dispatchEvent(new WheelEvent('wheel', { bubbles: true, cancelable: true, deltaY: 100 }))`)
  await delay(50)
  const wheelForward = await state()
  assert(shape(wheelForward, 'Click fade').visibility === 'visible', 'Wheel down must advance the next build.')
  await delay(350)
  await evaluate(`document.querySelector('.stage-shell').dispatchEvent(new WheelEvent('wheel', { bubbles: true, cancelable: true, deltaY: -100 }))`)
  await delay(30)
  const wheelBack = await state()
  assert(shape(wheelBack, 'Click fade').visibility === 'hidden', 'Wheel up must undo the current build.')

  await evaluate(`window.__decklineMediaCalls = []; window.__decklineMediaPauses = []; window.__decklineOriginalPlay = HTMLMediaElement.prototype.play; window.__decklineOriginalPause = HTMLMediaElement.prototype.pause; HTMLMediaElement.prototype.play = function () { window.__decklineMediaCalls.push({ id: this.dataset.mediaId, currentTime: this.currentTime, volume: this.volume, muted: this.muted }); return window.__decklineBlockNextMediaPlay ? (delete window.__decklineBlockNextMediaPlay, Promise.reject(new DOMException('Autoplay blocked', 'NotAllowedError'))) : Promise.resolve() }; HTMLMediaElement.prototype.pause = function () { window.__decklineMediaPauses.push(this.dataset.mediaId); return window.__decklineOriginalPause.apply(this, arguments) }; window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '5' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(300)
  const embeddedMedia = await state()
  assert(embeddedMedia.slide?.startsWith('15 / 67') && embeddedMedia.media.videos.length === 1 && embeddedMedia.media.audios.length === 3 && embeddedMedia.slideTransition?.clipKeyframes?.length === 5 && embeddedMedia.slideTransition.clipPath.startsWith('path('), `Embedded media and the OOXML dissolve transition must render on their source slide: ${JSON.stringify({ slide: embeddedMedia.slide, media: embeddedMedia.media, transition: embeddedMedia.slideTransition })}`)
  assert(embeddedMedia.media.videos[0].src.startsWith('blob:') && embeddedMedia.media.videos[0].poster.startsWith('blob:') && embeddedMedia.media.videos[0].controls && !embeddedMedia.media.videos[0].autoplay && !embeddedMedia.media.videos[0].muted, 'Embedded video must use local media and poster blobs with native controls and no forced playback.')
  const automaticMediaCalls = await evaluate(`window.__decklineMediaCalls`)
  assert(automaticMediaCalls.some(call => call.id === '3' && Math.abs(call.volume - 0.8) < 0.01 && !call.muted) && automaticMediaCalls.some(call => call.id === '2' && call.volume === 1 && !call.muted), `Timed p:audio/p:video nodes must start their respective targets using authored volume and mute settings: ${JSON.stringify(automaticMediaCalls)}`)
  const automaticMediaPauses = await evaluate(`window.__decklineMediaPauses`)
  assert(automaticMediaPauses.includes('2'), `A finite 100 ms media time node must stop its video target at its declared end: ${JSON.stringify(automaticMediaPauses)}`)
  assert(shape(embeddedMedia, 'Media after-effect').visibility === 'hidden', 'An afterEffect animation must wait for an indefinite timed media node to end.')
  assert(embeddedMedia.warnings.some(warning => warning.includes('Cross-slide video playback is not supported')), 'Unsupported cross-slide video must produce a visible warning.')
  await evaluate(`document.querySelector('.slide-audio').dispatchEvent(new Event('ended'))`)
  await delay(220)
  const mediaAfterEffectState = await state()
  assert(shape(mediaAfterEffectState, 'Media after-effect').visibility === 'visible', 'An indefinite media ended event must release the following afterEffect animation after its finite sibling duration.')
  await evaluate(`(() => { const video = document.querySelector('.slide-video'); Object.defineProperty(video, 'duration', { configurable: true, value: 1 }); Object.defineProperty(video, 'currentTime', { configurable: true, writable: true, value: 0.1 }); video.dispatchEvent(new Event('loadedmetadata')) })()`)
  await delay(20)
  const videoFade = await state()
  assert(Math.abs(videoFade.media.videos[0].volume - 0.5) < 0.02 && !videoFade.media.videos[0].muted && videoFade.media.videos[0].opacity === '1', `Video media fade must control its audio volume while leaving its picture visible: ${JSON.stringify(videoFade.media.videos[0])}`)
  assert(embeddedMedia.media.audios[0].src.startsWith('blob:') && embeddedMedia.media.audios[0].controls && !embeddedMedia.media.audios[0].autoplay && embeddedMedia.media.audios[0].readyState >= 1 && Math.abs(embeddedMedia.media.audios[0].duration - 0.5) < 0.02 && Math.abs(embeddedMedia.media.audios[0].currentTime - 0.1) < 0.02 && embeddedMedia.media.audios[0].muted, `Embedded WAV metadata, 100 ms trim, and fade-in start must load from a local blob: ${JSON.stringify(embeddedMedia.media.audios[0])}`)
  assert(embeddedMedia.media.audios[1].readyState >= 1 && Math.abs(embeddedMedia.media.audios[1].duration - 0.25) < 0.02 && embeddedMedia.media.warnings.some(warning => warning.name === 'Invalid trim' && warning.message.includes('full media is used')) && embeddedMedia.media.warnings.some(warning => warning.name === 'Invalid trim' && warning.message.includes('fades are ignored')), `Trim and fade ranges longer than their WAV must be reported and ignored: ${JSON.stringify(embeddedMedia.media)}`)
  await evaluate(`(() => { const audio = document.querySelector('.slide-audio'); audio.currentTime = 0.2; audio.dispatchEvent(new Event('seeking')) })()`)
  await delay(30)
  const fadingIn = await state()
  assert(Math.abs(fadingIn.media.audios[0].volume - 0.4) < 0.02 && !fadingIn.media.audios[0].muted, `Audio fade-in must scale the authored 80% timeline volume across its 200 ms range: ${JSON.stringify(fadingIn.media.audios[0])}`)
  await evaluate(`(() => { const audio = document.querySelector('.slide-audio'); audio.currentTime = 0.35; audio.dispatchEvent(new Event('seeking')) })()`)
  await delay(30)
  const fadingOut = await state()
  assert(Math.abs(fadingOut.media.audios[0].volume - 0.4) < 0.02, `Audio fade-out must scale the authored 80% timeline volume across its 100 ms range: ${JSON.stringify(fadingOut.media.audios[0])}`)
  await evaluate(`(() => { const audio = document.querySelector('.slide-audio'); audio.currentTime = 0.45; audio.dispatchEvent(new Event('seeking')); audio.dispatchEvent(new Event('timeupdate')) })()`)
  await delay(40)
  const trimmedMedia = await state()
  assert(Math.abs(trimmedMedia.media.audios[0].currentTime - 0.4) < 0.02 && trimmedMedia.media.audios[0].muted && trimmedMedia.media.warnings.some(warning => warning.name === 'Invalid trim'), `Seeking beyond the 100 ms end trim must clamp playback and finish the fade-out: ${JSON.stringify(trimmedMedia.media)}`)
  assert(!embeddedMedia.warnings.some(warning => warning.includes('PowerPoint autoplay/timing is not applied')), 'The old warning for all timed media playback must be removed after supported timeline media playback is enabled.')
  await evaluate(`(() => { const video = document.querySelector('.slide-video'); video.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true })); video.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, cancelable: true })) })()`)
  assert((await state()).slide?.startsWith('15 / 67'), 'Video clicks and media-focused navigation keys must not advance the slide.')

  const shapeMasksInitial = await state()
  const shapeMaskNames = ['Circle mask animation', 'Diamond mask animation', 'Box mask animation', 'Dissolve mask animation', 'Wheel mask animation', 'Wedge mask animation', 'Plus in mask animation', 'Plus out mask animation', 'Random horizontal bars animation', 'Random vertical bars animation', 'Strips downLeft animation', 'Strips upLeft animation', 'Strips downRight animation', 'Strips upRight animation']
  const shapeMaskStarts = shapeMaskNames.map(name => shape(shapeMasksInitial, name).clipPath)
  assert(shapeMaskStarts.every(clipPath => clipPath.startsWith('inset(') || clipPath.startsWith('path(')), `Circle, diamond, and box effects must have valid initial masks: ${JSON.stringify(shapeMaskStarts)}`)
  assert((shape(shapeMasksInitial, 'Wheel mask animation').clipPath.match(/M(?=\s*[-\d])/g) || []).length === 4 && (shape(shapeMasksInitial, 'Wedge mask animation').clipPath.match(/M(?=\s*[-\d])/g) || []).length === 1, 'The wheel and wedge filters must retain their authored spoke counts.')
  for (const [index, name] of shapeMaskNames.entries()) {
    await evaluate(`document.querySelector('.stage-shell').click()`)
    await delay(60)
    const active = await state()
    assert(shape(active, name).clipPath !== shapeMaskStarts[index], `${name} must interpolate after its click trigger.`)
    if (name === 'Random horizontal bars animation') assert(shape(active, name).clipPath.startsWith('path(') && (shape(active, name).clipPath.match(/M(?=\s*[-\d])/g) || []).length === 8, 'Horizontal random bars must render one progressive mask per horizontal band.')
    if (name === 'Random vertical bars animation') assert(shape(active, name).clipPath.startsWith('path(') && (shape(active, name).clipPath.match(/M(?=\s*[-\d])/g) || []).length === 9, 'Vertical random bars must render one progressive mask per vertical band.')
    if (name.startsWith('Strips ')) assert(shape(active, name).clipPath.startsWith('path(') && (shape(active, name).clipPath.match(/M(?=\s*[-\d])/g) || []).length === 8, `${name} must reveal through eight staggered bands.`)
    await delay(420)
  }
  const shapeMasksComplete = await state()
  assert(shape(shapeMasksComplete, 'Circle mask animation').clipPath.startsWith('path(') && shape(shapeMasksComplete, 'Diamond mask animation').clipPath.startsWith('inset(50%') && shape(shapeMasksComplete, 'Box mask animation').clipPath.startsWith('inset(0') && shape(shapeMasksComplete, 'Dissolve mask animation').clipPath.startsWith('inset(0') && (shape(shapeMasksComplete, 'Wheel mask animation').clipPath.match(/M(?=\s*[-\d])/g) || []).length === 4 && (shape(shapeMasksComplete, 'Wedge mask animation').clipPath.match(/M(?=\s*[-\d])/g) || []).length === 1 && shape(shapeMasksComplete, 'Plus in mask animation').clipPath.startsWith('inset(0') && shape(shapeMasksComplete, 'Plus out mask animation').clipPath.startsWith('inset(50%') && shape(shapeMasksComplete, 'Random horizontal bars animation').clipPath.startsWith('path(') && (shape(shapeMasksComplete, 'Random horizontal bars animation').clipPath.match(/M(?=\s*[-\d])/g) || []).length === 8 && shape(shapeMasksComplete, 'Random vertical bars animation').clipPath.startsWith('path(') && (shape(shapeMasksComplete, 'Random vertical bars animation').clipPath.match(/M(?=\s*[-\d])/g) || []).length === 9, `Circle, diamond, box, dissolve, wheel, wedge, plus, and random-bar masks must reach their OOXML final states: ${JSON.stringify(shapeMaskNames.map(name => [name, shape(shapeMasksComplete, name).clipPath]))}`)
  assert(shapeMaskNames.filter(name => name.startsWith('Strips ')).every(name => shape(shapeMasksComplete, name).clipPath.startsWith('path(') && (shape(shapeMasksComplete, name).clipPath.match(/M(?=\s*[-\d])/g) || []).length === 8), 'All four object strips filters must finish with eight visible bands.')

  const barnNames = ['inHorizontal', 'outHorizontal', 'inVertical', 'outVertical'].map(direction => `Barn ${direction} animation`)
  const barnStarts = barnNames.map(name => shape(shapeMasksComplete, name).clipPath)
  for (const [index, name] of barnNames.entries()) {
    await evaluate(`document.querySelector('.stage-shell').click()`)
    await delay(60)
    const active = await state()
    assert(shape(active, name).clipPath !== barnStarts[index] && shape(active, name).clipPath.startsWith('inset('), `${name} must animate through a CSS barn-door mask.`)
    await delay(420)
  }
  const barnsComplete = await state()
  assert(shape(barnsComplete, 'Barn inHorizontal animation').clipPath.startsWith('inset(0') && shape(barnsComplete, 'Barn inVertical animation').clipPath.startsWith('inset(0') && shape(barnsComplete, 'Barn outHorizontal animation').visibility === 'hidden' && shape(barnsComplete, 'Barn outVertical animation').visibility === 'hidden', 'Barn masks must finish their horizontal/vertical reveal or exit states.')
  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(60)
  const characterWipeState = await state()
  await evaluate(`window.__decklineBlockNextMediaPlay = true; document.querySelector('.stage-shell').click()`)
  await delay(20)
  const playFromCalls = await evaluate(`window.__decklineMediaCalls`)
  const lastPlayFromCall = playFromCalls.at(-1)
  const playFromState = await state()
  assert(lastPlayFromCall?.id === '3' && Math.abs(playFromState.media.audios[0].currentTime - 0.2) < 0.02 && playFromState.media.warnings.some(warning => warning.message.includes('browser blocked automatic playback')), `The click-timed PlayFrom command must seek the target and surface browser autoplay rejection: ${JSON.stringify({ lastPlayFromCall, currentTime: playFromState.media.audios[0].currentTime, warnings: playFromState.media.warnings })}`)
  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(20)
  const togglePauseCalls = await evaluate(`window.__decklineMediaCalls`)
  assert(togglePauseCalls.length === playFromCalls.length + 1 && togglePauseCalls.at(-1)?.id === '3', `The togglePause command must call play for a paused target: ${JSON.stringify(togglePauseCalls)}`)
  await evaluate(`document.querySelector('.slide-audio').dispatchEvent(new Event('play')); document.querySelector('.stage-shell').click()`)
  await delay(20)
  const pauseCalls = await evaluate(`window.__decklineMediaCalls`)
  assert(pauseCalls.length === togglePauseCalls.length, `The pause command must not start media playback: ${JSON.stringify(pauseCalls)}`)
  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(20)
  const stoppedMedia = await state()
  assert(Math.abs(stoppedMedia.media.audios[0].currentTime - 0.1) < 0.02, `The stop command must pause and return media to its trimmed start: ${JSON.stringify(stoppedMedia.media.audios[0])}`)
  await evaluate(`HTMLMediaElement.prototype.play = window.__decklineOriginalPlay; HTMLMediaElement.prototype.pause = window.__decklineOriginalPause; delete window.__decklineOriginalPlay; delete window.__decklineOriginalPause; delete window.__decklineMediaCalls; delete window.__decklineMediaPauses`)

  await evaluate(`document.querySelector('[data-tool=pen]').click(); (() => { const layer = document.querySelector('.ink-layer'); const rect = layer.getBoundingClientRect(); const event = (type, x, y, buttons) => layer.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, pointerId: 21, pointerType: 'mouse', clientX: rect.left + x, clientY: rect.top + y, buttons })); event('pointerdown', 30, 30, 1); event('pointermove', 90, 80, 1); event('pointerup', 90, 80, 0) })()`)
  await delay(20)
  await evaluate(`document.querySelector('[data-tool=highlighter]').click(); (() => { const layer = document.querySelector('.ink-layer'); const rect = layer.getBoundingClientRect(); const event = (type, x, y, buttons) => layer.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, pointerId: 22, pointerType: 'mouse', clientX: rect.left + x, clientY: rect.top + y, buttons })); event('pointerdown', 40, 100, 1); event('pointermove', 120, 100, 1); event('pointerup', 120, 100, 0) })()`)
  await delay(20)
  const inkDrawn = await state()
  assert(inkDrawn.ink.lines.length === 2 && inkDrawn.ink.lines.some(line => line.opacity === '0.4'), `Pen and highlighter strokes must persist on the current slide: ${JSON.stringify({ slide: inkDrawn.slide, ink: inkDrawn.ink })}`)
  await evaluate(`document.querySelector('[data-tool=laser]').click(); (() => { const layer = document.querySelector('.ink-layer'); const rect = layer.getBoundingClientRect(); layer.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, pointerId: 23, pointerType: 'mouse', clientX: rect.left + 70, clientY: rect.top + 70, buttons: 1 })) })()`)
  await delay(20)
  const laserActive = await state()
  assert(laserActive.ink.laserCount === 1, 'The laser tool must render a transient pointer.')
  await evaluate(`document.querySelector('.ink-layer').dispatchEvent(new PointerEvent('pointerup', { bubbles: true, cancelable: true, pointerId: 23, pointerType: 'mouse' })); document.querySelector('[data-tool=clear]').click(); document.querySelector('[aria-label="Zoom in"]').click()`)
  await delay(20)
  const clearedAndZoomed = await state()
  assert(clearedAndZoomed.ink.lines.length === 0 && clearedAndZoomed.ink.laserCount === 0, 'Clear must remove current-slide ink and pointer release must hide the laser.')
  assert(clearedAndZoomed.zoomTransform.startsWith('matrix(1.25'), 'The zoom control must scale the slide to 125%.')
  await evaluate(`document.querySelector('[data-tool=pointer]').click(); (() => { const stage = document.querySelector('.stage-shell'); const event = (type, id, x) => stage.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, pointerId: id, pointerType: 'touch', clientX: x, clientY: 150, buttons: type === 'pointerup' ? 0 : 1 })); event('pointerdown', 31, 100); event('pointerdown', 32, 200); event('pointermove', 32, 250); event('pointerup', 32, 250); event('pointerup', 31, 100) })()`)
  await delay(20)
  const pinched = await state()
  assert(pinched.zoomTransform.startsWith('matrix(1.88'), 'Two touch pointers must increase zoom from 125% to about 188%.')
  await delay(450)
  await evaluate(`document.querySelector('.stage-shell').click(); window.__crossSlideMedia = document.querySelector('[data-media-id="5"]')`)
  await delay(150)
  const crossSlideStart = await evaluate(`({ paused: window.__crossSlideMedia.paused, currentTime: window.__crossSlideMedia.currentTime, duration: window.__crossSlideMedia.duration })`)
  assert(!crossSlideStart.paused && crossSlideStart.currentTime > 0 && Math.abs(crossSlideStart.duration - 2) < 0.02, `The timed numSld media action must start its real embedded audio: ${JSON.stringify(crossSlideStart)}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }))`)
  await delay(650)
  await evaluate(`window.__crossSlideMedia = document.querySelector('.continuing-media-host [data-media-id="5"]')`)
  const crossSlideContinuation = await evaluate(`({ connected: window.__crossSlideMedia?.isConnected, paused: window.__crossSlideMedia?.paused, currentTime: window.__crossSlideMedia?.currentTime })`)
  assert(crossSlideContinuation.connected && !crossSlideContinuation.paused && crossSlideContinuation.currentTime > crossSlideStart.currentTime + 0.2, `numSld=2 audio must keep playing after its source slide is removed: ${JSON.stringify({ start: crossSlideStart, continuation: crossSlideContinuation })}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }))`)
  await delay(40)
  const crossSlideLimit = await evaluate(`({ paused: window.__crossSlideMedia.paused, currentTime: window.__crossSlideMedia.currentTime })`)
  assert(crossSlideLimit.paused, `numSld=2 audio must stop when leaving the second slide: ${JSON.stringify(crossSlideLimit)}`)
  await evaluate(`delete window.__crossSlideMedia`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '6' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const randomBarTransition = await state()
  const randomBarFrames = randomBarTransition.slideTransition?.clipKeyframes || []
  assert(randomBarTransition.slide?.startsWith('16 / 67') && randomBarFrames.length === 5 && randomBarFrames.every(frame => frame.startsWith('path(')) && (randomBarFrames[2].match(/M(?=\s*[-\d])/g) || []).length === 8 && Number(randomBarTransition.slideTransition.duration) === 750, `The horizontal random-bar slide transition must animate eight stable, interpolable masks: ${JSON.stringify(randomBarTransition.slideTransition)}`)
  await delay(800)
  const jumpToPresetSlide = slideNumber => evaluate(`(() => { for (const key of '${slideNumber}') window.dispatchEvent(new KeyboardEvent('keydown', { key })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' })) })()`)
  const presetCases = ['fallOver', 'drape', 'curtains', 'wind', 'prestige', 'fracture', 'crush', 'peelOff', 'pageCurlDouble', 'pageCurlSingle', 'airplane', 'origami']
  for (const [index, presetName] of presetCases.entries()) {
    const slideNumber = 54 + index
    await jumpToPresetSlide(slideNumber)
    await delay(80)
    const result = await state()
    const snapshot = await evaluate(`(() => { const underlay = document.querySelector('.slide-transition-underlay'); const outgoing = underlay?.firstElementChild; const animation = outgoing?.getAnimations()[0]; return { preset: document.querySelector('.stage-shell > .slide-host')?.dataset.transitionPreset, className: document.querySelector('.stage-shell')?.className, layer: getComputedStyle(underlay).zIndex, duration: animation?.effect.getTiming().duration, transforms: animation?.effect.getKeyframes().map(frame => frame.transform), origins: animation?.effect.getKeyframes().map(frame => frame.transformOrigin) } })()`)
    assert(result.slide?.startsWith(`${slideNumber} / 67`) && result.transitionEffect === 'preset' && result.transitionPreset === presetName && snapshot.preset === presetName && snapshot.className.includes('is-preset-transition') && Number(snapshot.layer) > 1 && Number(snapshot.duration) === 850 && snapshot.transforms?.length > 1, `p15:prstTrans ${presetName} must parse, animate the outgoing snapshot, and honor p14:dur: ${JSON.stringify({ slide: result.slide, effect: result.transitionEffect, preset: snapshot.preset, duration: snapshot.duration, snapshot, warnings: result.warnings })}`)
    if (presetName === 'fallOver') assert(snapshot.transforms.at(-1).includes('rotateX(-72deg)'), `p15:invY must mirror fallOver's vertical direction: ${JSON.stringify(snapshot.transforms)}`)
    if (presetName === 'peelOff') assert(snapshot.origins.at(-1) === 'right center' && snapshot.transforms.at(-1).includes('rotateY(-105deg)'), `p15:invX must mirror peelOff's hinge direction: ${JSON.stringify({ origins: snapshot.origins, transforms: snapshot.transforms })}`)
    await delay(850)
    assert(!(await evaluate(`document.querySelector('.stage-shell')?.classList.contains('is-preset-transition')`)) && !(await evaluate(`document.querySelector('.slide-transition-underlay')?.firstElementChild`)), `p15:prstTrans ${presetName} must clear its snapshot and class after the duration.`)
  }
  await jumpToPresetSlide(66)
  await delay(80)
  const invalidPreset = await state()
  assert(invalidPreset.slide?.startsWith('66 / 67') && invalidPreset.transitionEffect === 'fade' && invalidPreset.warnings.some(warning => warning.includes('notARealPreset') && warning.includes('falls back to fade')), `An unknown p15 preset must be warned and use a fade fallback: ${JSON.stringify({ slide: invalidPreset.slide, effect: invalidPreset.transitionEffect, warnings: invalidPreset.warnings })}`)
  await evaluate(`document.querySelector('.stage-shell')?.click()`)
  await delay(35)
  const visibilityMiddle = await state()
  assert(shape(visibilityMiddle, 'Visibility from-to').visibility === 'visible' && shape(visibilityMiddle, 'Visibility keyframes').visibility === 'hidden', `Discrete style.visibility must apply its active keyframe while keeping from visible: ${JSON.stringify({ fromTo: shape(visibilityMiddle, 'Visibility from-to'), keyframes: shape(visibilityMiddle, 'Visibility keyframes'), warnings: visibilityMiddle.warnings })}`)
  await delay(300)
  const visibilityEnd = await state()
  assert(shape(visibilityEnd, 'Visibility from-to').visibility === 'hidden' && shape(visibilityEnd, 'Visibility keyframes').visibility === 'visible', `Discrete style.visibility must hold its direct and keyframed end values: ${JSON.stringify({ fromTo: shape(visibilityEnd, 'Visibility from-to'), keyframes: shape(visibilityEnd, 'Visibility keyframes') })}`)
  await delay(550)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '7' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  await evaluate(`new Promise((resolve, reject) => { const started = performance.now(); const check = () => { const frame = document.querySelector('.stage-shell > .slide-host .slide-frame'); if (frame && getComputedStyle(frame).backgroundSize === '2px 8px') resolve(true); else if (performance.now() - started > 3000) reject(new Error('The flipped background tile did not finish rendering.')); else setTimeout(check, 20) }; check() })`)
  const verticalRandomBarTransition = await state()
  const verticalRandomBarFrames = verticalRandomBarTransition.slideTransition?.clipKeyframes || []
  assert(verticalRandomBarTransition.slide?.startsWith('17 / 67') && verticalRandomBarFrames.length === 5 && verticalRandomBarFrames.every(frame => frame.startsWith('path(')) && (verticalRandomBarFrames[2].match(/M(?=\s*[-\d])/g) || []).length === 15, `The vertical random-bar slide transition must animate fifteen stable, interpolable masks: ${JSON.stringify(verticalRandomBarTransition.slideTransition)}`)
  const tileSize = verticalRandomBarTransition.slideBackground?.size.match(/^([\d.]+)px ([\d.]+)px$/)?.slice(1).map(Number)
  const tileOffsets = [...(verticalRandomBarTransition.slideBackground?.position.matchAll(/calc\(100% ([+-]) ([\d.]+)px\)/g) || [])]
  const expectedTileOffsets = [914400, -457200].map((offset, index) => offset * (verticalRandomBarTransition.slideBackground?.frameWidth || 0) / 12192000 + (index ? 4 : 1))
  assert(/^url\(["']?blob:/.test(verticalRandomBarTransition.slideBackground?.image || '') && verticalRandomBarTransition.slideBackground?.repeat === 'repeat' && tileSize?.length === 2 && Math.abs(tileSize[0] - 2) < 0.01 && Math.abs(tileSize[1] - 8) < 0.01 && tileOffsets.length === 2 && tileOffsets.every((match, index) => Math.abs((match[1] === '-' ? -1 : 1) * Number(match[2]) - expectedTileOffsets[index]) < 0.1) && !verticalRandomBarTransition.warnings.some(warning => warning.includes('Tiled slide background')), `DrawingML tile scale, bottom-right alignment, xy-flip, and EMU offsets must render with a local image and without fallback warnings: ${JSON.stringify({ background: verticalRandomBarTransition.slideBackground, tileOffsets: tileOffsets.map(match => match[0]), expectedTileOffsets, warnings: verticalRandomBarTransition.warnings })}`)
  const tileFlipPixels = await evaluate(`(async () => { const frame = document.querySelector('.stage-shell > .slide-host .slide-frame'); const backgroundImage = getComputedStyle(frame).backgroundImage; const url = backgroundImage.startsWith('url("') ? backgroundImage.slice(5, -2) : backgroundImage.slice(4, -1); let response; let blob; try { response = await fetch(url); blob = await response.blob(); const bitmap = await createImageBitmap(blob); const canvas = document.createElement('canvas'); canvas.width = bitmap.width; canvas.height = bitmap.height; const context = canvas.getContext('2d'); context.drawImage(bitmap, 0, 0); const pixels = [[0,0],[1,0],[2,0],[3,0],[0,1],[1,1],[2,1],[3,1],[0,2],[1,2],[0,3],[1,3]].map(([x,y]) => [...context.getImageData(x,y,1,1).data].slice(0,3).map(value => value.toString(16).padStart(2,'0')).join('')); return { width: bitmap.width, height: bitmap.height, pixels } } catch (error) { return { url, status: response?.status, type: blob?.type, size: blob?.size, error: String(error) } } })()`)
  assert(tileFlipPixels.width === 4 && tileFlipPixels.height === 4 && tileFlipPixels.pixels.join(',') === 'ff0000,00ff00,00ff00,ff0000,0000ff,ffff00,ffff00,0000ff,0000ff,ffff00,ff0000,00ff00', `Tile flip=xy must mirror alternating rows and columns in the generated pattern: ${JSON.stringify(tileFlipPixels)}`)
  const smartArtCachedRoot = verticalRandomBarTransition.shapes.find(element => element.name === 'SmartArt cached drawing fixture: Root node')
  const smartArtCachedChild = verticalRandomBarTransition.shapes.find(element => element.name === 'SmartArt cached drawing fixture: Child node')
  const smartArtRunText = verticalRandomBarTransition.textRuns.map(run => run.text)
  const smartArtFallback = verticalRandomBarTransition.shapes.find(element => element.name === 'SmartArt hierarchy fallback fixture')
  const smartArtFallbackParagraphs = verticalRandomBarTransition.paragraphs.filter(paragraph => ['Fallback root', 'Fallback child'].includes(paragraph.text.trim()))
  assert(smartArtCachedRoot?.visibility === 'visible' && smartArtCachedChild?.visibility === 'visible' && smartArtCachedRoot.frameRect.width > 0 && smartArtCachedChild.frameRect.width > 0 && smartArtCachedChild.frameRect.left > smartArtCachedRoot.frameRect.left && smartArtCachedChild.frameRect.top > smartArtCachedRoot.frameRect.top && smartArtRunText.includes('Cached diagram root') && smartArtRunText.includes('Cached diagram child') && verticalRandomBarTransition.warnings.some(warning => warning.includes('last saved diagram drawing')) && smartArtFallback?.visibility === 'visible' && smartArtFallbackParagraphs.length === 2 && Number.parseFloat(smartArtFallbackParagraphs[1].marginLeft) > Number.parseFloat(smartArtFallbackParagraphs[0].marginLeft) && verticalRandomBarTransition.warnings.some(warning => warning.includes('SmartArt is shown as an indented text list')), `SmartArt must render its cached diagram geometry inside the frame and retain a parent-indented text fallback when the cache is absent: ${JSON.stringify({ cached: [smartArtCachedRoot, smartArtCachedChild], fallback: smartArtFallback, paragraphs: smartArtFallbackParagraphs, warnings: verticalRandomBarTransition.warnings })}`)
  const olePreview = verticalRandomBarTransition.shapes.find(element => element.name === 'OLE preview fixture')
  assert(olePreview?.kind === 'picture' && olePreview.imageSource?.startsWith('blob:') && olePreview.visibility === 'visible' && olePreview.frameRect.width > 0 && olePreview.frameRect.height > 0 && verticalRandomBarTransition.warnings.some(warning => warning.includes('OLE object is displayed from its embedded preview image')), `OLE must render its local preview in the graphic-frame bounds without opening the object: ${JSON.stringify({ preview: olePreview, warnings: verticalRandomBarTransition.warnings })}`)
  const oleNoPreview = verticalRandomBarTransition.shapes.find(element => element.name === 'OLE no-preview fixture')
  assert(oleNoPreview?.kind === 'shape' && oleNoPreview.visibility === 'visible' && oleNoPreview.frameRect.width > 0 && oleNoPreview.frameRect.height > 0 && verticalRandomBarTransition.textRuns.some(run => run.text === 'Word.Document.12') && verticalRandomBarTransition.warnings.some(warning => warning.includes('no usable embedded preview image')), `An OLE object without a preview must retain its frame, display its ProgID label, and report the fallback: ${JSON.stringify({ placeholder: oleNoPreview, text: verticalRandomBarTransition.textRuns.map(run => run.text), warnings: verticalRandomBarTransition.warnings })}`)
  const inheritedParagraph = paragraph(verticalRandomBarTransition, '→Inherited placeholder')
  const inheritedRun = verticalRandomBarTransition.textRuns.find(run => run.text === 'Inherited placeholder')
  assert(shape(verticalRandomBarTransition, 'Inherited paragraph').boxShadow.includes('rgba(0, 0, 0, 0.5)'), `An unoverridden slide placeholder must inherit its layout outer shadow: ${shape(verticalRandomBarTransition, 'Inherited paragraph').boxShadow}`)
  assert(shape(verticalRandomBarTransition, 'Inherited paragraph').borderStyle === 'dashed', `An unoverridden slide placeholder must inherit its layout preset line dash: ${shape(verticalRandomBarTransition, 'Inherited paragraph').borderStyle}`)
  assert(verticalRandomBarTransition.textFrames.find(frame => frame.name === 'Inherited paragraph')?.whiteSpace === 'nowrap', 'An unoverridden slide placeholder must inherit wrap="none" from its layout body properties.')
  const inheritedVerticalText = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'Inherited paragraph')
  assert(inheritedVerticalText?.writingMode === 'vertical-rl' && inheritedVerticalText.textOrientation === 'sideways', 'A slide placeholder must inherit vertical text flow and orientation from its layout body properties.')
  const vertical270Frame = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'Vertical 270 text')
  assert(vertical270Frame?.writingMode === 'vertical-lr' && vertical270Frame.textOrientation === 'sideways' && vertical270Frame.direction === 'rtl', 'vert270 text must flow upward in columns from left to right.')
  const eastAsianFrame = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'East Asian vertical text')
  assert(eastAsianFrame?.writingMode === 'vertical-rl' && eastAsianFrame.textOrientation === 'upright', 'eaVert text must use upright glyphs in right-to-left vertical columns.')
  assert(verticalRandomBarTransition.textFrames.find(frame => frame.name === 'Explicit horizontal text')?.writingMode === 'horizontal-tb', 'An explicit horz value must remain horizontal.')
  const twoColumnFrame = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'Two-column text')
  const expectedColumnGap = 457200 * (twoColumnFrame?.frameWidth || 0) / 12192000
  const columnXs = (twoColumnFrame?.runRects || []).map(rect => rect.left)
  assert(twoColumnFrame?.columnCount === '2' && Math.abs(Number.parseFloat(twoColumnFrame.columnGap) - expectedColumnGap) < 0.1 && columnXs.length > 1 && Math.max(...columnXs) - Math.min(...columnXs) > expectedColumnGap, `a:bodyPr numCol and spcCol must create equal text columns with the specified EMU gap and flow text into the next column: ${JSON.stringify({ twoColumnFrame, expectedColumnGap })}`)
  const ltrColumnFlow = await evaluate(`(() => {
    const node = document.querySelector('.stage-shell > .slide-host .slide-element[title="Two-column text"] .text-run')?.firstChild
    if (!node || node.nodeType !== Node.TEXT_NODE) return undefined
    const rectAt = (start, end) => { const range = document.createRange(); range.setStart(node, start); range.setEnd(node, end); const rect = range.getBoundingClientRect(); return { left: rect.left, right: rect.right } }
    return { first: rectAt(0, 5), last: rectAt(Math.max(0, node.textContent.length - 5), node.textContent.length) }
  })()`)
  assert(twoColumnFrame?.direction === 'ltr' && twoColumnFrame.paragraphDirection === 'ltr' && ltrColumnFlow?.first.left < ltrColumnFlow?.last.left, `a:bodyPr rtlCol="0" must preserve left-to-right column overflow and paragraph direction: ${JSON.stringify({ frame: twoColumnFrame && { direction: twoColumnFrame.direction, paragraphDirection: twoColumnFrame.paragraphDirection }, flow: ltrColumnFlow })}`)
  const rtlColumnFrame = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'Two-column RTL text')
  const rtlColumnFlow = await evaluate(`(() => {
    const node = document.querySelector('.stage-shell > .slide-host .slide-element[title="Two-column RTL text"] .text-run')?.firstChild
    if (!node || node.nodeType !== Node.TEXT_NODE) return undefined
    const rectAt = (start, end) => { const range = document.createRange(); range.setStart(node, start); range.setEnd(node, end); const rect = range.getBoundingClientRect(); return { left: rect.left, right: rect.right } }
    return { first: rectAt(0, 5), last: rectAt(Math.max(0, node.textContent.length - 5), node.textContent.length) }
  })()`)
  assert(rtlColumnFrame?.columnCount === '2' && rtlColumnFrame.direction === 'rtl' && rtlColumnFrame.paragraphDirection === 'ltr' && rtlColumnFrame.runRects?.length > 1 && rtlColumnFlow?.first.left > rtlColumnFlow?.last.left, `a:bodyPr rtlCol must start overflow in the rightmost column without changing paragraph text direction: ${JSON.stringify({ frame: rtlColumnFrame, flow: rtlColumnFlow })}`)
  const rtlParagraphFrame = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'RTL paragraph direction')
  assert(rtlParagraphFrame?.direction === 'rtl' && rtlParagraphFrame.paragraphDirection === 'rtl', `a:pPr rtl must control paragraph direction independently from a:bodyPr rtlCol: ${JSON.stringify(rtlParagraphFrame)}`)
  const wordArtFrame = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'WordArt vertical')
  assert(wordArtFrame?.writingMode === 'vertical-rl' && wordArtFrame.textOrientation === 'upright'
    && verticalRandomBarTransition.warnings.some(warning => warning.includes('wordArtVert') && warning.includes('CSS writing modes')),
    'wordArtVert text must use upright vertical CSS flow and report its shaping approximation.')
  const mongolianFrame = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'Mongolian vertical')
  assert(mongolianFrame?.writingMode === 'vertical-lr' && mongolianFrame.textOrientation === 'upright'
    && verticalRandomBarTransition.warnings.some(warning => warning.includes('mongolianVert') && warning.includes('CSS writing modes')),
    'mongolianVert text must use left-to-right upright vertical CSS flow and report its shaping approximation.')
  const wordArtRtlFrame = verticalRandomBarTransition.textFrames.find(frame => frame.name === 'WordArt vertical RTL')
  assert(wordArtRtlFrame?.writingMode === 'vertical-rl' && wordArtRtlFrame.textOrientation === 'upright'
    && wordArtRtlFrame.paragraphRects?.length === 2 && wordArtRtlFrame.paragraphRects[0].left > wordArtRtlFrame.paragraphRects[1].left
    && !verticalRandomBarTransition.warnings.some(warning => warning.includes('wordArtVertRtl')),
    `wordArtVertRtl must stack glyphs top-to-bottom and flow paragraphs right-to-left without a fallback warning: ${JSON.stringify({ wordArtRtlFrame, warnings: verticalRandomBarTransition.warnings })}`)
  assert(inheritedParagraph?.textAlign === 'center' && inheritedParagraph.lineHeight === '1.25' && inheritedParagraph.marginLeft && inheritedParagraph.marker === '→', `Placeholder paragraph defaults must cascade from master through layout: ${JSON.stringify({ inheritedParagraph, warnings: verticalRandomBarTransition.warnings })}`)
  const inheritedFontSizeRatio = Number.parseFloat(inheritedRun?.fontSize || '0') / verticalRandomBarTransition.hostFrame.width
  assert(Math.abs(inheritedFontSizeRatio - 32 * 12_700 / 12_192_000) < 0.00001 && inheritedRun?.fontWeight === '700' && inheritedRun.fontStyle === 'italic' && inheritedRun.color === 'rgb(255, 0, 0)' && inheritedRun.textDecorationLine.includes('line-through'), `Placeholder run defaults must preserve layout font size, master emphasis/color, and strike: ${JSON.stringify({ inheritedRun, inheritedFontSizeRatio })}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '8' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const strips = await state()
  const stripsFrames = strips.slideTransition?.clipKeyframes || []
  const stripsMiddle = stripsFrames[2] || ''
  assert(strips.slide?.startsWith('18 / 67') && stripsFrames.length === 5 && stripsFrames.every(frame => frame.startsWith('path(')) && (stripsMiddle.match(/M(?=\s*[-\d])/g) || []).length === 8 && Number(strips.slideTransition.duration) === 750 && !strips.warnings.some(warning => warning.includes('strips')), `The four-corner strips transition must render eight stable staggered masks without falling back: ${JSON.stringify(strips.slideTransition)}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '9' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const plusTransition = await state()
  const plusFrames = plusTransition.slideTransition?.clipKeyframes || []
  assert(plusTransition.slide?.startsWith('19 / 67') && plusFrames.length === 5 && plusFrames.every(frame => frame.startsWith('path("M')) && new Set(plusFrames).size === 5 && !plusTransition.warnings.some(warning => warning.includes('plus')), `The plus transition must expand a stable cross mask without falling back: ${JSON.stringify(plusTransition.slideTransition)}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '0' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const wedgeTransition = await state()
  const wedgeFrames = wedgeTransition.slideTransition?.clipKeyframes || []
  assert(wedgeTransition.slide?.startsWith('20 / 67') && wedgeFrames.length === 5 && wedgeFrames.every(frame => frame.startsWith('path("M')) && wedgeFrames.every(frame => (frame.match(/L/g) || []).length === 13) && !wedgeTransition.warnings.some(warning => warning.includes('wedge')), `The wedge transition must sweep a single radial sector without falling back: ${JSON.stringify(wedgeTransition.slideTransition)}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const newsflash = await state()
  assert(newsflash.slide?.startsWith('21 / 67') && newsflash.slideTransition?.keyframes[0]?.includes('rotate(360deg)') && newsflash.slideTransition.keyframes[0]?.includes('scale(0.2)') && newsflash.slideTransition.keyframes.at(-1)?.includes('rotate(0deg)') && Number(newsflash.slideTransition.opacityKeyframes[0]) === 0 && !newsflash.warnings.some(warning => warning.includes('newsflash')), `Newsflash must grow and counter-rotate the incoming slide without falling back: ${JSON.stringify(newsflash.slideTransition)}`)
  await delay(800)
  await evaluate(`window.__decklineRandomValues = [0, 0.9999]; window.__decklineOriginalRandom = Math.random; Math.random = () => window.__decklineRandomValues.shift() ?? 0.5; window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const randomFade = await state()
  assert(randomFade.slide?.startsWith('22 / 67') && Number(randomFade.slideTransition?.opacityKeyframes?.[0]) === 0 && Number(randomFade.slideTransition?.opacityKeyframes.at(-1)) === 1, `The first random-transition draw must select fade: ${JSON.stringify(randomFade.slideTransition)}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const randomPush = await state()
  await evaluate(`Math.random = window.__decklineOriginalRandom; delete window.__decklineOriginalRandom; delete window.__decklineRandomValues`)
  assert(randomPush.slide?.startsWith('22 / 67') && randomPush.slideTransition?.keyframes[0]?.includes('translateX(100%)') && randomPush.slideTransition.keyframes.at(-1)?.includes('translate(0px, 0px)'), `The next random-transition draw must select a different supported effect: ${JSON.stringify(randomPush.slideTransition)}`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(800)
  await evaluate(`window.__decklineRandomValues = [0.675]; window.__decklineOriginalRandom = Math.random; Math.random = () => window.__decklineRandomValues.shift() ?? 0.5; window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const randomWheel = await state()
  const randomWheelFrames = randomWheel.slideTransition?.clipKeyframes || []
  assert(randomWheel.slide?.startsWith('22 / 67') && randomWheelFrames.length === 5 && (randomWheelFrames[2]?.match(/M/g) || []).length === 4, `The random wheel transition must use the schema default of four spokes: ${JSON.stringify(randomWheel.slideTransition)}`)
  await evaluate(`Math.random = window.__decklineOriginalRandom; delete window.__decklineOriginalRandom; delete window.__decklineRandomValues`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(800)
  await evaluate(`window.__decklineRandomValues = [0.925]; window.__decklineOriginalRandom = Math.random; Math.random = () => window.__decklineRandomValues.shift() ?? 0.5; window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const randomFlash = await state()
  assert(randomFlash.slide?.startsWith('22 / 67') && randomFlash.flashTransition?.opacityKeyframes?.map(Number).join(',') === '0,1,1,0' && Number(randomFlash.flashTransition.duration) === 750, `A random transition draw must be able to select flash: ${JSON.stringify(randomFlash.flashTransition)}`)
  await evaluate(`Math.random = window.__decklineOriginalRandom; delete window.__decklineOriginalRandom; delete window.__decklineRandomValues`)
  await delay(800)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const fadeThroughBlack = await state()
  assert(fadeThroughBlack.slide?.startsWith('23 / 67') && fadeThroughBlack.slideTransition?.opacityKeyframes?.map(Number).join(',') === '0,0,1' && fadeThroughBlack.transitionUnderlay.opacityKeyframes?.map(Number).join(',') === '1,0,0' && Number(fadeThroughBlack.slideTransition.duration) === 1000 && fadeThroughBlack.stageClasses.includes('is-fading-through-black') && !fadeThroughBlack.warnings.some(warning => warning.includes('through-black')), `Fade through black must fade the outgoing and incoming slides around a black midpoint: ${JSON.stringify({ incoming: fadeThroughBlack.slideTransition, outgoing: fadeThroughBlack.transitionUnderlay, classes: fadeThroughBlack.stageClasses, warnings: fadeThroughBlack.warnings })}`)
  await delay(1_050)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const longTransition = await state()
  const verticalCombFrames = longTransition.slideTransition?.clipKeyframes || []
  assert(longTransition.slide?.startsWith('24 / 67') && verticalCombFrames.length === 5 && verticalCombFrames.every(frame => frame.startsWith('path(')) && (verticalCombFrames[2].match(/M(?=\s*[-\d])/g) || []).length === 15 && Number(longTransition.slideTransition.duration) === 45_000 && !longTransition.warnings.some(warning => warning.includes('comb')), `Vertical comb must render animated bars and preserve explicit p14:dur: ${JSON.stringify(longTransition.slideTransition)}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '5' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(40)
  const fastTransition = await state()
  const horizontalCombFrames = fastTransition.slideTransition?.clipKeyframes || []
  assert(fastTransition.slide?.startsWith('25 / 67') && horizontalCombFrames.length === 5 && horizontalCombFrames.every(frame => frame.startsWith('path(')) && (horizontalCombFrames[2].match(/M(?=\s*[-\d])/g) || []).length === 8 && Number(fastTransition.slideTransition.duration) === 500 && !fastTransition.warnings.some(warning => warning.includes('comb')), `Horizontal comb must render animated bars and preserve fast transition timing: ${JSON.stringify(fastTransition.slideTransition)}`)
  const autoFitFrames = fastTransition.textFrames.filter(frame => frame.name?.includes('auto-fit calibration') || frame.name === 'Auto fit calibration')
  assert(autoFitFrames.length === 2 && autoFitFrames.every(frame => frame.scrollHeight <= frame.clientHeight), `Auto-fit text at large and small font sizes must fit its frame: ${JSON.stringify(autoFitFrames)}`)
  assert(autoFitFrames.every(frame => frame.borderStyle === 'none' && frame.borderWidth === '0px'), `A 1pt noFill line must not render as a CSS border: ${JSON.stringify(autoFitFrames)}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '6' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const flash = await state()
  assert(flash.slide?.startsWith('26 / 67') && flash.flashTransition?.opacityKeyframes?.map(Number).join(',') === '0,1,1,0' && Number(flash.flashTransition.duration) === 900 && flash.flashTransition.backgroundColor === 'rgb(255, 255, 255)' && !flash.warnings.some(warning => warning.includes('flash')), `p14:flash must be selected from its AlternateContent choice and render its white overlay using p14:dur: ${JSON.stringify({ transition: flash.flashTransition, warnings: flash.warnings })}`)
  await delay(950)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '7' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const doors = await state()
  assert(doors.slide?.startsWith('27 / 67') && doors.slideTransition?.clipKeyframes?.length === 2 && doors.slideTransition.clipKeyframes[0]?.includes('50%') && Number(doors.slideTransition.duration) === 600 && !doors.warnings.some(warning => warning.includes('doors')), `p14:doors must select its AlternateContent choice, open from the vertical center seam, and preserve p14:dur: ${JSON.stringify({ transition: doors.slideTransition, warnings: doors.warnings })}`)
  await delay(650)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '8' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(60)
  const windowTransition = await state()
  assert(windowTransition.slide?.startsWith('28 / 67') && windowTransition.slideTransition?.clipKeyframes?.[0]?.includes('50%') && Number(windowTransition.slideTransition.duration) === 700 && !windowTransition.warnings.some(warning => warning.includes('window')), `p14:window must select its AlternateContent choice and expand from a centered aperture: ${JSON.stringify({ transition: windowTransition.slideTransition, warnings: windowTransition.warnings })}`)
  await delay(750)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '9' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const prism = await state()
  assert(prism.slide?.startsWith('29 / 67') && prism.stageClasses.includes('is-prism-transition') && Number(prism.slideTransition?.duration) === 850 && prism.slideTransition?.keyframes[0]?.includes('rotate3d(0, 1, 0, 90deg)') && prism.transitionUnderlay.keyframes?.at(-1)?.includes('rotate3d(0, 1, 0, -90deg)') && prism.warnings.some(warning => warning.includes('isContent flag requests separate background and content planes')) && !prism.warnings.some(warning => warning.includes('prism slide transition falls back')), `p14:prism must animate the incoming and outgoing faces, apply isInverted, preserve p14:dur, and report the unsupported isContent split: ${JSON.stringify({ incoming: prism.slideTransition, outgoing: prism.transitionUnderlay, classes: prism.stageClasses, warnings: prism.warnings })}`)
  await delay(900)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '0' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const flip = await state()
  assert(flip.slide?.startsWith('30 / 67') && flip.stageClasses.includes('is-flip-transition') && Number(flip.slideTransition?.duration) === 800 && flip.slideTransition?.keyframes[0]?.includes('rotate3d(0, 1, 0, 180deg)') && flip.transitionUnderlay.keyframes?.at(-1)?.includes('rotate3d(0, 1, 0, -180deg)') && !flip.warnings.some(warning => warning.includes('flip slide transition falls back')), `p14:flip must honor rightward direction and p14:dur while rotating incoming and outgoing slides in 3D: ${JSON.stringify({ incoming: flip.slideTransition, outgoing: flip.transitionUnderlay, classes: flip.stageClasses, warnings: flip.warnings })}`)
  await delay(850)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const reveal = await state()
  assert(reveal.slide?.startsWith('31 / 67') && reveal.stageClasses.includes('is-fading-through-black') && Number(reveal.slideTransition?.duration) === 1050 && reveal.slideTransition?.keyframes[0]?.includes('translateX(100%)') && reveal.slideTransition?.keyframes[1]?.includes('translateX(100%)') && reveal.slideTransition?.opacityKeyframes?.map(Number).join(',') === '0,0,1' && reveal.transitionUnderlay.keyframes?.at(-1)?.includes('translateX(-100%)') && reveal.transitionUnderlay.opacityKeyframes?.map(Number).join(',') === '1,0,0' && !reveal.warnings.some(warning => warning.includes('reveal slide transition falls back')), `p14:reveal must move left through a black midpoint and honor its direction, thruBlk, and p14:dur: ${JSON.stringify({ incoming: reveal.slideTransition, outgoing: reveal.transitionUnderlay, classes: reveal.stageClasses, warnings: reveal.warnings })}`)
  await delay(1100)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const switchTransition = await state()
  assert(switchTransition.slide?.startsWith('32 / 67') && switchTransition.stageClasses.includes('is-switch-transition') && Number(switchTransition.slideTransition?.duration) === 950 && switchTransition.slideTransition?.keyframes[0]?.includes('rotate3d(0, 1, 0, 90deg)') && switchTransition.transitionUnderlay.keyframes?.at(-1)?.includes('rotate3d(0, 1, 0, -90deg)') && !switchTransition.warnings.some(warning => warning.includes('switch slide transition falls back')), `p14:switch must honor rightward direction and p14:dur while rotating between slide planes: ${JSON.stringify({ incoming: switchTransition.slideTransition, outgoing: switchTransition.transitionUnderlay, classes: switchTransition.stageClasses, warnings: switchTransition.warnings })}`)
  await delay(1000)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const pan = await state()
  assert(pan.slide?.startsWith('33 / 67') && Number(pan.slideTransition?.duration) === 900 && pan.slideTransition?.keyframes[0]?.includes('translate(100%, -100%)') && pan.transitionUnderlay.keyframes?.at(-1)?.includes('translate(-100%, 100%)') && !pan.warnings.some(warning => warning.includes('pan slide transition falls back')), `p14:pan must pair both slide planes in its left/down direction and honor p14:dur: ${JSON.stringify({ incoming: pan.slideTransition, outgoing: pan.transitionUnderlay, warnings: pan.warnings })}`)
  await delay(950)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const ripple = await state()
  const rippleMovingPath = ripple.slideTransition?.clipPath
  await delay(60)
  const rippleAdvancedPath = (await state()).slideTransition?.clipPath
  const rippleStartPoint = ripple.slideTransition?.clipKeyframes?.[0]?.match(/^path\("M\s*(-?[\d.]+)\s+(-?[\d.]+)/)
  assert(ripple.slide?.startsWith('34 / 67') && Number(ripple.slideTransition?.duration) === 1200 && ripple.slideTransition.clipKeyframes?.length === 6 && rippleStartPoint && Number(rippleStartPoint[1]) > 0 && Number(rippleStartPoint[2]) === 0 && rippleMovingPath !== rippleAdvancedPath && ripple.slideTransition.clipKeyframes.at(-1)?.startsWith('path("M') && !ripple.warnings.some(warning => warning.includes('ripple')), `p14:ripple direction, timing, and path keyframes: ${JSON.stringify({ slide: ripple.slide, duration: ripple.slideTransition?.duration, frameCount: ripple.slideTransition?.clipKeyframes?.length, start: rippleStartPoint, last: ripple.slideTransition?.clipKeyframes?.at(-1), warnings: ripple.warnings })}`)
  await delay(1250)
  const tileStats = (path) => {
    const moves = path?.match(/\bM\b/g)?.length || 0
    const lines = path?.match(/\bL\b/g)?.length || 0
    return { tiles: moves, pointsPerTile: lines / Math.max(1, moves) }
  }
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '5' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const glitterDiamondLeft = await state()
  const glitterDiamondPath = glitterDiamondLeft.slideTransition?.clipKeyframes?.[2] || ''
  const glitterDiamondStats = tileStats(glitterDiamondPath)
  const glitterDiamondMovingPath = glitterDiamondLeft.slideTransition?.clipPath
  await delay(60)
  const glitterDiamondAdvancedPath = (await state()).slideTransition?.clipPath
  assert(glitterDiamondLeft.slide?.startsWith('35 / 67') && Number(glitterDiamondLeft.slideTransition?.duration) === 1000 && glitterDiamondStats.tiles > 0 && glitterDiamondStats.pointsPerTile === 4 && glitterDiamondMovingPath !== glitterDiamondAdvancedPath && !glitterDiamondLeft.warnings.some(warning => warning.includes('glitter')), 'p14:glitter diamond must use four-point tiles, honor its left direction and duration, and interpolate.')
  await delay(900)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '6' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const glitterDiamondRightDown = await state()
  const glitterRightDownPath = glitterDiamondRightDown.slideTransition?.clipKeyframes?.[2]
  assert(glitterDiamondRightDown.slide?.startsWith('36 / 67') && Number(glitterDiamondRightDown.slideTransition?.duration) === 1100 && glitterDiamondRightDown.slideTransition?.clipKeyframes?.[2] !== glitterDiamondPath && !glitterDiamondRightDown.warnings.some(warning => warning.includes('glitter')), 'p14:glitter must order diamond tiles from the opposite corner when dir=rd.')
  await delay(1050)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '7' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const glitterHexagon = await state()
  const glitterHexagonPath = glitterHexagon.slideTransition?.clipKeyframes?.[2] || ''
  const glitterHexagonStats = tileStats(glitterHexagonPath)
  const glitterHexagonMovingPath = glitterHexagon.slideTransition?.clipPath
  await delay(60)
  const glitterHexagonAdvancedPath = (await state()).slideTransition?.clipPath
  assert(glitterHexagon.slide?.startsWith('37 / 67') && Number(glitterHexagon.slideTransition?.duration) === 1200 && glitterHexagonStats.tiles > 0 && glitterHexagonStats.pointsPerTile === 6 && glitterHexagonMovingPath !== glitterHexagonAdvancedPath && !glitterHexagon.warnings.some(warning => warning.includes('glitter')), 'p14:glitter hexagon must use six-point tiles, preserve its down direction and duration, and interpolate.')
  await delay(1150)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '8' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const shredStrip = await state()
  const shredStripPath = shredStrip.slideTransition?.clipKeyframes?.[2] || ''
  const shredStripStats = tileStats(shredStripPath)
  const shredStripMovingPath = shredStrip.slideTransition?.clipPath
  await delay(60)
  const shredStripAdvanced = await state()
  const shredStripAdvancedPath = shredStripAdvanced.slideTransition?.clipPath
  assert(shredStrip.slide?.startsWith('38 / 67') && Number(shredStrip.slideTransition?.duration) === 1000 && shredStrip.slideTransition?.clipKeyframes?.length === 5 && shredStripStats.tiles === 16 && shredStripStats.pointsPerTile === 3 && shredStripMovingPath !== shredStripAdvancedPath && !shredStrip.warnings.some(warning => warning.includes('shred')), `p14:shred defaults to incoming vertical strips and honors p14:dur: ${JSON.stringify({ slide: shredStrip.slide, effect: shredStrip.transitionEffect, stageClasses: shredStrip.stageClasses, underlay: shredStrip.transitionUnderlay, transition: shredStrip.slideTransition, advancedTransition: shredStripAdvanced.slideTransition, shredStripStats, warnings: shredStrip.warnings, moving: shredStripMovingPath !== shredStripAdvancedPath })}`)
  await delay(900)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '9' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const shredRectangleOut = await evaluate(`(() => {
    const underlay = document.querySelector('.slide-transition-underlay')
    const outgoing = underlay?.querySelector('.slide-host')
    const frames = outgoing?.getAnimations()[0]?.effect.getKeyframes().map(frame => frame.clipPath) || []
    const stats = path => ({ tiles: (path?.match(/\\bM\\b/g) || []).length, pointsPerTile: (path?.match(/\\bL\\b/g) || []).length / Math.max(1, (path?.match(/\\bM\\b/g) || []).length) })
    return { slide: document.querySelector('.slide-position')?.textContent?.replace(/\\s+/g, ' ').trim(), className: underlay?.className, duration: outgoing?.getAnimations()[0]?.effect.getTiming().duration, start: stats(frames[0]), changed: frames[0] !== frames.at(-1), frameCount: frames.length }
  })()`)
  assert(shredRectangleOut.slide?.startsWith('39 / 67') && Number(shredRectangleOut.duration) === 1000 && shredRectangleOut.frameCount === 5 && shredRectangleOut.start.tiles === 70 && shredRectangleOut.start.pointsPerTile === 3 && shredRectangleOut.changed && shredRectangleOut.className.includes('is-shredding-out'), `p14:shred rectangle out must animate the outgoing snapshot as 70 tiles: ${JSON.stringify(shredRectangleOut)}`)
  await delay(1000)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '0' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const flythroughIn = await state()
  const flythroughInOutgoing = await evaluate(`(() => { const underlay = document.querySelector('.slide-transition-underlay'); const host = underlay?.querySelector('.slide-host'); const animation = host?.getAnimations()[0]; return { className: underlay?.className, duration: animation?.effect.getTiming().duration, transforms: animation?.effect.getKeyframes().map(frame => frame.transform) } })()`)
  assert(flythroughIn.slide?.startsWith('40 / 67') && Number(flythroughIn.slideTransition?.duration) === 900 && flythroughIn.slideTransition?.keyframes?.[0]?.includes('scale(2.4)') && Number(flythroughIn.slideTransition?.opacityKeyframes?.[0]) === 0 && flythroughIn.slideTransition?.keyframes?.at(-1)?.includes('scale(1)') && flythroughInOutgoing.className.includes('is-flythrough') && Number(flythroughInOutgoing.duration) === 900 && flythroughInOutgoing.transforms?.at(-1)?.includes('scale(0.35)') && !flythroughIn.warnings.some(warning => warning.includes('flythrough')), `p14:flythrough defaults to an unbounced inward movement of both slide planes: ${JSON.stringify({ slide: flythroughIn.slide, duration: flythroughIn.slideTransition?.duration, transforms: flythroughIn.slideTransition?.keyframes, opacity: flythroughIn.slideTransition?.opacityKeyframes, warnings: flythroughIn.warnings, outgoing: flythroughInOutgoing })}`)
  await delay(830)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const flythroughBounceOut = await state()
  const flythroughOutOutgoing = await evaluate(`(() => { const host = document.querySelector('.slide-transition-underlay > .slide-host'); const animation = host?.getAnimations()[0]; return { duration: animation?.effect.getTiming().duration, transforms: animation?.effect.getKeyframes().map(frame => frame.transform) } })()`)
  assert(flythroughBounceOut.slide?.startsWith('41 / 67') && Number(flythroughBounceOut.slideTransition?.duration) === 900 && flythroughBounceOut.slideTransition?.keyframes?.length === 3 && flythroughBounceOut.slideTransition.keyframes[0]?.includes('scale(0.35)') && flythroughBounceOut.slideTransition.keyframes[1]?.includes('scale(1.08)') && Number(flythroughBounceOut.slideTransition.opacityKeyframes?.[0]) === 0 && Number(flythroughOutOutgoing.duration) === 900 && flythroughOutOutgoing.transforms?.at(-1)?.includes('scale(2.4)') && !flythroughBounceOut.warnings.some(warning => warning.includes('flythrough')), `p14:flythrough out honors hasBounce and animates the outgoing snapshot in front of the incoming slide: ${JSON.stringify({ slide: flythroughBounceOut.slide, duration: flythroughBounceOut.slideTransition?.duration, transforms: flythroughBounceOut.slideTransition?.keyframes, opacity: flythroughBounceOut.slideTransition?.opacityKeyframes, warnings: flythroughBounceOut.warnings, outgoing: flythroughOutOutgoing })}`)
  await delay(850)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const warpOut = await state()
  const warpOutSnapshot = await evaluate(`(() => { const underlay = document.querySelector('.slide-transition-underlay'); const host = underlay?.querySelector('.slide-host'); const animation = host?.getAnimations()[0]; return { className: underlay?.className, layerZIndex: getComputedStyle(underlay).zIndex, hostZIndex: getComputedStyle(document.querySelector('.stage-shell > .slide-host')).zIndex, duration: animation?.effect.getTiming().duration, transforms: animation?.effect.getKeyframes().map(frame => frame.transform), opacity: animation?.effect.getKeyframes().map(frame => frame.opacity) } })()`)
  assert(warpOut.slide?.startsWith('42 / 67') && Number(warpOut.slideTransition?.duration) === 850 && warpOutSnapshot.className.includes('is-warping-out') && Number(warpOutSnapshot.layerZIndex) > Number(warpOutSnapshot.hostZIndex) && warpOutSnapshot.transforms?.[1]?.includes('perspective(1200px)') && warpOutSnapshot.transforms?.[1]?.includes('rotateY(-24deg)') && warpOutSnapshot.opacity?.join(',') === '1,0' && !warpOut.warnings.some(warning => warning.includes('warp')), `p14:warp defaults to out, animates the outgoing snapshot, and honors p14:dur: ${JSON.stringify({ slide: warpOut.slide, incoming: warpOut.slideTransition, outgoing: warpOutSnapshot, warnings: warpOut.warnings })}`)
  await delay(850)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '3' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const warpIn = await state()
  assert(warpIn.slide?.startsWith('43 / 67') && Number(warpIn.slideTransition?.duration) === 850 && warpIn.slideTransition?.keyframes?.[0]?.includes('rotateY(24deg)') && warpIn.slideTransition?.keyframes?.at(-1)?.includes('scale(1)') && warpIn.slideTransition.opacityKeyframes?.join(',') === '0,1,1' && !warpIn.warnings.some(warning => warning.includes('warp')), `p14:warp honors dir=in by warping the incoming slide: ${JSON.stringify({ slide: warpIn.slide, transition: warpIn.slideTransition, warnings: warpIn.warnings })}`)
  await delay(850)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const vortexDefault = await state()
  const vortexDefaultSnapshot = await evaluate(`(() => { const underlay = document.querySelector('.slide-transition-underlay'); const host = underlay?.querySelector('.slide-host'); const animation = host?.getAnimations()[0]; return { className: underlay?.className, layerZIndex: getComputedStyle(underlay).zIndex, hostZIndex: getComputedStyle(document.querySelector('.stage-shell > .slide-host')).zIndex, duration: animation?.effect.getTiming().duration, transforms: animation?.effect.getKeyframes().map(frame => frame.transform), opacity: animation?.effect.getKeyframes().map(frame => frame.opacity) } })()`)
  assert(vortexDefault.slide?.startsWith('44 / 67') && Number(vortexDefaultSnapshot.duration) === 850 && vortexDefaultSnapshot.className.includes('is-vortexing') && Number(vortexDefaultSnapshot.layerZIndex) > Number(vortexDefaultSnapshot.hostZIndex) && vortexDefaultSnapshot.transforms?.at(-1)?.includes('rotate3d(0, 0, 1, -540deg)') && vortexDefaultSnapshot.opacity?.join(',') === '1,0' && !vortexDefault.warnings.some(warning => warning.includes('vortex')), `p14:vortex defaults to dir=l and rotates the outgoing slide: ${JSON.stringify({ slide: vortexDefault.slide, snapshot: vortexDefaultSnapshot, warnings: vortexDefault.warnings })}`)
  await delay(850)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '5' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(80)
  const vortexRight = await state()
  assert(vortexRight.slide?.startsWith('45 / 67') && vortexRight.transitionUnderlay.keyframes?.at(-1)?.includes('rotate3d(0, 0, 1, 540deg)') && vortexRight.transitionUnderlay.opacityKeyframes?.join(',') === '1,0' && Number(vortexRight.transitionUnderlay.duration) === 850 && !vortexRight.warnings.some(warning => warning.includes('vortex')), `p14:vortex honors dir=r and p14:dur: ${JSON.stringify({ slide: vortexRight.slide, snapshot: vortexRight.transitionUnderlay, warnings: vortexRight.warnings })}`)
  await delay(850)
  const jumpToFixtureSlide = slideNumber => evaluate(`(() => { for (const key of '${slideNumber}') window.dispatchEvent(new KeyboardEvent('keydown', { key })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' })) })()`)
  const verifySideTransition = async (slideNumber, effect, className, incoming, outgoing) => {
    await jumpToFixtureSlide(slideNumber)
    await delay(80)
    const result = await state()
    assert(result.slide?.startsWith(`${slideNumber} / 67`) && result.stageClasses.includes(className) && Number(result.slideTransition?.duration) === 850 && result.slideTransition?.keyframes?.[0]?.includes(incoming) && result.transitionUnderlay.keyframes?.at(-1)?.includes(outgoing) && !result.warnings.some(warning => warning.includes(`${effect} slide transition falls back`)), `p14:${effect} direction and p14:dur: ${JSON.stringify({ slide: result.slide, classes: result.stageClasses, incoming: result.slideTransition, outgoing: result.transitionUnderlay, warnings: result.warnings })}`)
    await delay(850)
  }
  await verifySideTransition(46, 'ferris', 'is-ferris-transition', 'rotateZ(180deg)', 'rotateZ(-180deg)')
  await verifySideTransition(47, 'ferris', 'is-ferris-transition', 'rotateZ(-180deg)', 'rotateZ(180deg)')
  await verifySideTransition(48, 'gallery', 'is-gallery-transition', 'translateX(36%)', 'translateX(-36%)')
  await verifySideTransition(49, 'gallery', 'is-gallery-transition', 'translateX(-36%)', 'translateX(36%)')
  await verifySideTransition(50, 'conveyor', 'is-conveyor-transition', 'translateX(100%)', 'translateX(-100%)')
  await verifySideTransition(51, 'conveyor', 'is-conveyor-transition', 'translateX(-100%)', 'translateX(100%)')
  await jumpToFixtureSlide(52)
  await delay(80)
  const morph = await state()
  const morphSnapshot = await evaluate(`(() => { const current = document.querySelector('.stage-shell > .slide-host'); const outgoing = document.querySelector('.slide-transition-underlay > .slide-host'); const incomingShape = current?.querySelector('.slide-element[data-element-name="Conveyor right"]'); const outgoingShape = outgoing?.querySelector('.slide-element[data-element-name="Conveyor right"]'); const animation = incomingShape?.getAnimations()[0]; return { className: document.querySelector('.stage-shell')?.className, duration: animation?.effect.getTiming().duration, left: animation?.effect.getKeyframes().map(frame => frame.left), top: animation?.effect.getKeyframes().map(frame => frame.top), width: animation?.effect.getKeyframes().map(frame => frame.width), oldOpacity: outgoingShape?.getAnimations()[0]?.effect.getKeyframes().map(frame => frame.opacity) } })()`)
  assert(morph.slide?.startsWith('52 / 67') && morphSnapshot.className.includes('is-morph-transition') && Number(morphSnapshot.duration) === 1000 && morphSnapshot.left?.[0] !== morphSnapshot.left?.at(-1) && morphSnapshot.top?.[0] !== morphSnapshot.top?.at(-1) && morphSnapshot.width?.[0] !== morphSnapshot.width?.at(-1) && morphSnapshot.oldOpacity?.join(',') === '1,0' && !morph.warnings.some(warning => warning.includes('Morph byObject')), `p159:morph byObject must animate a uniquely named matching shape and honor p14:dur: ${JSON.stringify({ slide: morph.slide, snapshot: morphSnapshot, warnings: morph.warnings })}`)
  await delay(1_000)
  assert(!(await evaluate(`document.querySelector('.stage-shell')?.classList.contains('is-morph-transition')`)) && !(await evaluate(`document.querySelector('.slide-transition-underlay')?.firstElementChild`)), 'The morph transition must clear its temporary slide snapshot when finished.')
  await jumpToFixtureSlide(53)
  await delay(80)
  const morphByWord = await state()
  assert(morphByWord.slide?.startsWith('53 / 67') && morphByWord.transitionEffect === 'morph' && Number(morphByWord.slideTransition?.duration) === 650 && morphByWord.slideTransition.opacityKeyframes?.join(',') === '0,1' && morphByWord.transitionUnderlay.opacityKeyframes?.join(',') === '1,0' && morphByWord.warnings.some(warning => warning.includes('Morph byWord matching is not rendered yet; the slide uses a crossfade.')), `p159:morph byWord must be explicitly warned and crossfade: ${JSON.stringify({ slide: morphByWord.slide, transition: morphByWord.slideTransition, outgoing: morphByWord.transitionUnderlay, warnings: morphByWord.warnings })}`)
  await delay(700)
  assert(!(await evaluate(`document.querySelector('.stage-shell')?.classList.contains('is-morph-transition')`)) && !(await evaluate(`document.querySelector('.slide-transition-underlay')?.firstElementChild`)), 'The byWord fallback crossfade must clear its temporary slide snapshot when finished.')
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '7' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(800)
  await evaluate(`document.querySelector('[data-tool=speaker]').click()`)
  const presenter = await evaluate(`(() => ({ visible: !!document.querySelector('[data-presenter-view]'), current: document.querySelector('.presenter-current-panel .presenter-panel-title')?.innerText, next: document.querySelector('.presenter-next-panel .presenter-panel-title')?.innerText, notes: document.querySelector('[data-presenter-notes]')?.textContent }))()`)
  assert(presenter.visible && presenter.current?.includes('17 / 67') && presenter.next?.includes('18 / 67') && presenter.notes?.includes('Introduce the main point.\nPause for questions.'), `Speaker view must show the current slide, next slide, and notes from the linked notes slide: ${JSON.stringify(presenter)}`)
  await delay(1_050)
  const presenterClock = await evaluate(`document.querySelector('.presenter-clock')?.textContent?.trim()`)
  const elapsedSeconds = presenterClock?.split(':').reduce((total, part) => total * 60 + Number(part), 0) || 0
  assert(elapsedSeconds >= 1, `The speaker view elapsed timer must advance: ${presenterClock}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }))`)
  const presenterNavigation = await evaluate(`document.querySelector('.presenter-current-panel .presenter-panel-title')?.innerText`)
  assert(presenterNavigation?.includes('18 / 67'), 'Arrow-key navigation must update the active slide while speaker view is open.')
  const projectorButtonPoint = await evaluate(`(() => { const rect = document.querySelector('[data-open-projector]')?.getBoundingClientRect(); return rect && { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()`)
  assert(projectorButtonPoint, 'Speaker view must expose a second-display button.')
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', ...projectorButtonPoint })
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', ...projectorButtonPoint, button: 'left', clickCount: 1 })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', ...projectorButtonPoint, button: 'left', clickCount: 1 })
  const projectorStatus = () => evaluate(`(() => { const status = document.querySelector('.presenter-display-status'); return status && { connected: status.dataset.projectorConnected, number: Number(status.dataset.projectorNumber), step: Number(status.dataset.projectorStep), blank: status.dataset.projectorBlank, triggerCount: Number(status.dataset.projectorTriggerCount) } })()`)
  let projector = null
  for (let attempt = 0; attempt < 50; attempt++) {
    projector = await projectorStatus()
    if (projector?.connected === 'true' && projector.number === 18) break
    await delay(100)
  }
  assert(projector?.connected === 'true' && projector.number === 18, `The separate projector window must load and acknowledge the active slide: ${JSON.stringify(projector)}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }))`)
  const speakerState = await evaluate(`(() => { const view = document.querySelector('[data-presenter-view]'); return { number: Number(view?.dataset.currentSlide), step: Number(view?.dataset.animationStep) } })()`)
  for (let attempt = 0; attempt < 30; attempt++) {
    projector = await projectorStatus()
    if (projector?.number === speakerState.number && Number(projector?.step) === speakerState.step) break
    await delay(80)
  }
  assert(projector?.number === speakerState.number && Number(projector?.step) === speakerState.step,
    `Projector navigation and animation-step state must follow the speaker controller: ${JSON.stringify({ speakerState, projector })}`)
  await jumpToFixtureSlide(44)
  await delay(100)
  const mirroredTransition = await evaluate(`(() => {
    const projectorWindow = globalThis.__decklineOpenedWindows?.find(candidate => !candidate.closed && candidate.document?.querySelector('[data-projector-window]'))
    const inspect = root => {
      const underlay = root?.querySelector('.slide-transition-underlay')
      const host = underlay?.querySelector(':scope > .slide-host')
      const animation = host?.getAnimations()[0]
      return {
        slide: root?.querySelector(':scope > .slide-host')?.getAttribute('data-slide-number'),
        effect: underlay?.className.includes('is-vortexing'),
        duration: animation?.effect.getTiming().duration,
        currentTime: animation?.currentTime,
      }
    }
    return {
      main: inspect(document.querySelector('.stage-shell')),
      speaker: inspect(document.querySelector('.presenter-transition-shell')),
      projector: inspect(projectorWindow?.document.querySelector('.projector-transition-shell')),
    }
  })()`)
  const mirroredProgress = Object.values(mirroredTransition).map(state => Number(state.currentTime))
  assert(Object.values(mirroredTransition).every(state => state.slide === '44' && state.effect && Number(state.duration) === 850)
    && mirroredProgress.every(Number.isFinite) && Math.max(...mirroredProgress) - Math.min(...mirroredProgress) < 120,
    `The main stage, Speaker View, and projector must run one synchronized slide transition: ${JSON.stringify(mirroredTransition)}`)
  await delay(900)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  for (let attempt = 0; attempt < 30; attempt++) {
    projector = await projectorStatus()
    if (projector?.number === 1 && Number(projector?.step) === 0) break
    await delay(80)
  }
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }))`)
  for (let attempt = 0; attempt < 30; attempt++) {
    projector = await projectorStatus()
    if (projector?.number === 1 && Number(projector?.step) === 1) break
    await delay(80)
  }
  const buildPreview = await evaluate(`(() => {
    const opacity = root => { const item = root?.querySelector('.slide-element[title="Click fade"]'); return item ? Number.parseFloat(getComputedStyle(item).opacity) : null }
    return { speakerStep: Number(document.querySelector('[data-presenter-view]')?.dataset.animationStep), projectorStep: Number(document.querySelector('.presenter-display-status')?.dataset.projectorStep), mainOpacity: opacity(document.querySelector('.stage-shell > .slide-host')), speakerOpacity: opacity(document.querySelector('.presenter-current-frame')) }
  })()`)
  assert(projector?.number === 1 && Number(projector?.step) === 1 && buildPreview.speakerStep === 1 && buildPreview.projectorStep === 1
    && buildPreview.mainOpacity !== null && buildPreview.speakerOpacity !== null && Math.abs(buildPreview.mainOpacity - buildPreview.speakerOpacity) < 0.2,
    `Speaker and projector previews must advance the same build with synchronized timing: ${JSON.stringify({ projector, buildPreview })}`)
  await evaluate(`document.querySelector('.stage-shell > .slide-host .slide-element[title="Scale emphasis"]')?.click()`)
  for (let attempt = 0; attempt < 40; attempt++) {
    projector = await projectorStatus()
    if (projector?.triggerCount === 1) break
    await delay(50)
  }
  assert(projector?.triggerCount === 1, `A shape-triggered animation must reach the projector state: ${JSON.stringify(projector)}`)
  await delay(400)
  const triggeredPreview = await evaluate(`(() => {
    const opacity = (root, getStyle = getComputedStyle) => {
      const element = root?.querySelector('.slide-element[title="Scale emphasis"]')
      return element ? Number.parseFloat(getStyle(element).opacity) : null
    }
    const projectorWindow = globalThis.__decklineOpenedWindows?.find(candidate => !candidate.closed && candidate.document?.querySelector('[data-projector-window]'))
    const projectorElement = projectorWindow?.document.querySelector('.projector-stage .slide-element[title="Scale emphasis"]')
    return {
      main: opacity(document.querySelector('.stage-shell > .slide-host')),
      speaker: opacity(document.querySelector('.presenter-current-frame')),
      projector: projectorElement ? Number.parseFloat(projectorWindow.getComputedStyle(projectorElement).opacity) : null,
      projectorSlide: Number(projectorWindow?.document.querySelector('.projector-transition-shell > .slide-host')?.dataset.slideNumber),
    }
  })()`)
  assert(triggeredPreview.projectorSlide === 1 && [triggeredPreview.main, triggeredPreview.speaker, triggeredPreview.projector].every(value => value !== null && Math.abs(value - 0.4) < 0.03),
    `The main stage, speaker preview, and projector must render the same completed shape-triggered animation: ${JSON.stringify({ projector, triggeredPreview })}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '1' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '8' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  for (let attempt = 0; attempt < 30; attempt++) {
    projector = await projectorStatus()
    if (projector?.number === 18 && Number(projector?.step) === 0) break
    await delay(80)
  }
  assert(projector?.number === 18 && Number(projector?.step) === 0, `Numeric slide navigation must synchronize back to slide 18: ${JSON.stringify(projector)}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'b' }))`)
  for (let attempt = 0; attempt < 25 && (await projectorStatus())?.blank !== 'black'; attempt++) await delay(80)
  projector = await projectorStatus()
  assert(projector?.blank === 'black', `Black-screen state must synchronize to the projector: ${JSON.stringify(projector)}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'b' }))`)
  for (let attempt = 0; attempt < 25 && (await projectorStatus())?.blank !== ''; attempt++) await delay(80)
  projector = await projectorStatus()
  assert(projector?.blank === '', `Clearing the screen blank must synchronize to the projector: ${JSON.stringify(projector)}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))`)
  assert(!await evaluate(`!!document.querySelector('[data-presenter-view]')`), 'Escape must close speaker view.')
  await evaluate(`document.querySelector('[data-tool=sorter]').click()`)
  const sorterBefore = await evaluate(`(() => ({ visible: !!document.querySelector('[data-sorter-view]'), cards: document.querySelectorAll('.sorter-card').length, range: document.querySelector('.sorter-header span')?.textContent, first: document.querySelector('.sorter-card .sorter-slide-button')?.getAttribute('aria-label'), second: document.querySelectorAll('.sorter-card .sorter-slide-button')[1]?.getAttribute('aria-label') }))()`)
  assert(sorterBefore.visible && sorterBefore.cards === 48 && sorterBefore.range?.includes('1–48 / 67'), `The slide sorter must show the first bounded group and total deck size: ${JSON.stringify(sorterBefore)}`)
  const sorterBeforeOrder = await evaluate(`([...document.querySelectorAll('.sorter-card .sorter-slide-button')].slice(0, 2).map(button => button.getAttribute('aria-label')))`)
  await evaluate(`(() => { const cards = [...document.querySelectorAll('.sorter-card')]; const transfer = new DataTransfer(); cards[0].dispatchEvent(new DragEvent('dragstart', { bubbles: true, dataTransfer: transfer })); cards[1].dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: transfer })) })()`)
  await delay(20)
  const sorterAfterOrder = await evaluate(`([...document.querySelectorAll('.sorter-card .sorter-slide-button')].slice(0, 2).map(button => button.getAttribute('aria-label')))`)
  const slideName = (label) => label?.split(': ').at(-1)
  assert(slideName(sorterAfterOrder[0]) === slideName(sorterBeforeOrder[1]) && slideName(sorterAfterOrder[1]) === slideName(sorterBeforeOrder[0]), `Dragging a slide over another must reorder them: ${JSON.stringify({ before: sorterBeforeOrder, after: sorterAfterOrder })}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', altKey: true }))`)
  const sorterKeyboard = await evaluate(`document.querySelector('.sorter-card.is-active .sorter-slide-button')?.getAttribute('aria-label')`)
  assert(sorterKeyboard === 'Go to slide 17: Slide 18', `Alt+arrow must move the selected slide while preserving selection: ${sorterKeyboard}`)
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))`)
  assert(!await evaluate(`!!document.querySelector('[data-sorter-view]')`), 'Escape must close slide sorter.')
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: '2' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: '4' })); window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))`)
  await delay(50)
  assert((await state()).slide?.startsWith('24 / 67'), 'The timed-advance fixture slide must be active.')
  await evaluate(`document.querySelector('.stage-shell').click()`)
  await delay(50)
  assert((await state()).slide?.startsWith('24 / 67'), 'A slide with advClick="0" must ignore stage clicks after its build steps.')
  await delay(5_100)
  assert((await state()).slide?.startsWith('25 / 67'), 'advTm must automatically advance to the next slide after its delay.')
  await evaluate(`window.__decklineFullscreenOriginal = HTMLElement.prototype.requestFullscreen; HTMLElement.prototype.requestFullscreen = function () { return Promise.resolve() }; window.dispatchEvent(new KeyboardEvent('keydown', { key: 'F5', shiftKey: true }))`)
  await delay(50)
  assert((await state()).slide?.startsWith('25 / 67'), 'Shift+F5 must present the current slide.')
  await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'F5' }))`)
  await delay(50)
  assert((await state()).slide?.startsWith('01 / 67'), 'F5 must start the presentation from its first slide.')
  await jumpToFixtureSlide(67)
  await delay(80)
  const textRangeSlide = await state()
  assert(!textRangeSlide.warnings.some(warning => warning.includes('Text-range motion animations are not rendered yet.')), 'Character-range motion must not be skipped by the parser.')
  const textRangeRuns = () => evaluate(`(() => [...(document.querySelector('.stage-shell > .slide-host')?.querySelectorAll('.slide-element[title="Character transform target"] .text-run') || [])].map(run => ({ text: run.textContent, start: Number(run.dataset.charStart), end: Number(run.dataset.charEnd), transform: getComputedStyle(run).transform, maskImage: getComputedStyle(run).maskImage, color: getComputedStyle(run).color, opacity: Number(getComputedStyle(run).opacity), display: getComputedStyle(run).display })))()`)
  const textParagraphMask = title => evaluate(`(() => { const target = document.querySelector('.stage-shell > .slide-host .slide-element[title="${title}"] .text-paragraph'); return target ? getComputedStyle(target).maskImage : 'missing' })()`)
  const textRangeBefore = await textRangeRuns()
  assert(textRangeSlide.slide?.startsWith('67 / 67') && textRangeBefore.length === 4 && textRangeBefore.map(run => run.text).join('') === 'PPTX' && !textRangeSlide.warnings.some(warning => warning.includes('unsupported target or text range was skipped') || warning.includes('unsupported target or property was skipped')), `Text-range scale and rotation targets must parse without a fallback warning: ${JSON.stringify({ slide: textRangeSlide.slide, runs: textRangeBefore, warnings: textRangeSlide.warnings })}`)
  await evaluate(`document.querySelector('[aria-label="Next animation step or slide"]')?.click()`)
  await delay(240)
  const textRangeAfter = await textRangeRuns()
  assert(textRangeAfter[0]?.opacity > 0.3 && textRangeAfter[0]?.opacity < 0.8
    && textRangeAfter[1]?.opacity > 0.3 && textRangeAfter[1]?.opacity < 0.8
    && textRangeAfter[2]?.opacity === 0.35 && textRangeAfter[3]?.opacity === 0.35,
    `Numeric opacity and p:set must apply only to their character ranges: ${JSON.stringify(textRangeAfter.map(run => ({ range: [run.start, run.end], opacity: run.opacity })))}`)
  const rangeMatrix = transform => transform.match(/^matrix\(([^)]+)\)$/)?.[1].split(',').map(Number)
  const scaleOnly = rangeMatrix(textRangeAfter[0]?.transform || '')
  const scaleAndRotation = rangeMatrix(textRangeAfter[1]?.transform || '')
  const presetScaleAndRotation = rangeMatrix(textRangeAfter[2]?.transform || '')
  const paragraphBlindsMask = await textParagraphMask('Paragraph blinds target')
  const paragraphCheckerMask = await textParagraphMask('Paragraph checker target')
  const textRangeMaskKinds = textRangeAfter.map(run => run.maskImage.includes('data:image/svg+xml') ? 'checker' : run.maskImage !== 'none' ? 'other' : 'none')
  assert((await state()).slide?.startsWith('67 / 67') && textRangeAfter.map(run => `${run.start}-${run.end}`).join(',') === '0-1,1-2,2-3,3-4'
    && scaleOnly && Math.hypot(scaleOnly[0], scaleOnly[1]) > 1.1 && Math.hypot(scaleOnly[2], scaleOnly[3]) > 1.05
    && scaleAndRotation && Math.abs(scaleAndRotation[1]) > 0.01 && presetScaleAndRotation && Math.abs(presetScaleAndRotation[1]) > 0.01
    && Math.hypot(presetScaleAndRotation[0], presetScaleAndRotation[1]) > 1.1
    && textRangeMaskKinds.join(',') === 'none,checker,checker,none'
    && paragraphBlindsMask.startsWith('repeating-linear-gradient(') && paragraphCheckerMask.includes('data:image/svg+xml')
    && !textRangeSlide.warnings.some(warning => warning.includes('Text-range slide') || warning.includes('character-range checkerboard'))
    && textRangeAfter[3]?.transform === 'none', `Paragraph and character mask animations plus p:charRg transforms must affect only their targets: ${JSON.stringify({ runRanges: textRangeAfter.map(run => [run.start, run.end]), textRangeMaskKinds, paragraphBlindsMask: paragraphBlindsMask.startsWith('repeating-linear-gradient('), paragraphCheckerMask: paragraphCheckerMask.includes('data:image/svg+xml'), warnings: textRangeSlide.warnings })}`)
  const motionOnFirstCharacter = rangeMatrix(textRangeAfter[0]?.transform || '')
  const motionOnAdjacentCharacters = textRangeAfter.slice(1, 3).map(run => rangeMatrix(run.transform))
  assert(motionOnFirstCharacter && Math.abs(motionOnFirstCharacter[4]) > 20 && Math.abs(motionOnFirstCharacter[5]) > 5
    && motionOnAdjacentCharacters.every(matrix => matrix && Math.abs(matrix[4]) < 0.01 && Math.abs(matrix[5]) < 0.01)
    && textRangeAfter[3]?.color !== textRangeBefore[3]?.color,
    `p:charRg motion must translate only the targeted text segment along its path: ${JSON.stringify(textRangeAfter)}`)
  await delay(1_000)
  const afterGroupRuns = await textRangeRuns()
  assert(afterGroupRuns[3]?.color === 'rgb(0, 0, 255)', `An afterGroup build must begin once the preceding animation group ends: ${JSON.stringify(afterGroupRuns)}`)
  await evaluate(`HTMLElement.prototype.requestFullscreen = window.__decklineFullscreenOriginal; delete window.__decklineFullscreenOriginal`)
  assert(runtimeErrors.length === 0, `Browser exceptions: ${runtimeErrors.join('; ')}`)

  const bottomThumbnailWindow = await evaluate(`(async () => {
    const list = document.querySelector('.thumb-list')
    const horizontal = getComputedStyle(list).flexDirection === 'row'
    if (horizontal) list.scrollLeft = list.scrollWidth
    else list.scrollTop = list.scrollHeight
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))
    const items = [...list.querySelectorAll('.thumb-button')]
    const result = { rendered: items.length, first: Number(items[0]?.dataset.thumbnailIndex), last: Number(items.at(-1)?.dataset.thumbnailIndex) }
    list.scrollTop = 0
    list.scrollLeft = 0
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))
    return result
  })()`)
  assert(bottomThumbnailWindow.rendered > 0 && bottomThumbnailWindow.rendered < 67 && bottomThumbnailWindow.last >= 63,
    `Scrolling to the final slide must virtualize around the tail of the deck: ${JSON.stringify(bottomThumbnailWindow)}`)
  await call('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true })
  const mobileThumbnailWindow = await evaluate(`(async () => {
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))
    const list = document.querySelector('.thumb-list')
    const direction = getComputedStyle(list).flexDirection
    list.scrollLeft = list.scrollWidth
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))
    const items = [...list.querySelectorAll('.thumb-button')]
    return { direction, rendered: items.length, last: Number(items.at(-1)?.dataset.thumbnailIndex), total: Number(document.querySelector('.slide-count')?.textContent) }
  })()`)
  await call('Emulation.clearDeviceMetricsOverride')
  assert(mobileThumbnailWindow.direction === 'row' && mobileThumbnailWindow.rendered > 0
    && mobileThumbnailWindow.rendered < mobileThumbnailWindow.total && mobileThumbnailWindow.last >= 63,
    `The mobile horizontal thumbnail strip must virtualize around the end of the deck: ${JSON.stringify(mobileThumbnailWindow)}`)

  console.log('Animation, media, presenter, and sorter verification passed: custom-geometry arcs, cached-data bar/line/pie/doughnut charts, timing, four-direction slide filters, character-range wipe and by-letter/by-word/by-shape iteration, linear/discrete font-size and opacity tavLst keyframes, numeric opacity attributes, text-range opacity and p:set, p:set visibility, fixed font-size, opacity, and colors, style/fill/stroke/shadow color and opacity animation, and line breaks, object blinds, checkerboard, random bars, four-direction object strips, four barn-door object filters, circle, diamond, box, plus, dissolve, wheel, wheelReverse, wedge and wipe effects, click/double-click/mouse-over/mouse-out shape triggers, fades, cut, blinds, checkerboard, comb, dissolve, wheel, wheelReverse, random-bar, strips, plus and wedge, newsflash, flash, doors, window, prism, directional p14 pan/ripple/glitter/shred/flythrough/warp/vortex/ferris/gallery/conveyor transitions, p159 morph byObject geometry matching and warned byWord fallback, random selection, cover, pull, split in/out, circle, diamond, honeycomb, paired push, zoom transitions, linked slide notes, current/next presenter view, speaker navigation, slide sorting by drag and keyboard, embedded media loading and controls, scale, rotation, RGB/HSL color, straight and cubic motion, tables, groups, controls, and ink.')
  }
} finally {
  await Promise.race([closeBrowser(), delay(1_000)])
  socket?.close()
  if (browser && browser.exitCode === null) {
    const exited = new Promise(resolve => browser.once('exit', resolve))
    browser.kill()
    await Promise.race([exited, delay(1_000)])
  }
  await delay(300)
  if (keepFixture) {
    await rm(profilePath, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 })
    console.log(`Retained fixture: ${fixturePath}`)
  } else {
    await rm(runtimeDir, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 })
  }
}
