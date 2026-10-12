import { _ as _plugin_vue_export_helper_default } from '../virtual/entry.mjs';
import { u as useSlideTransition, S as SlideStage_default, p as paragraphEntries, s as scriptSegments, c as cssFontFamily } from './useSlideTransition-DQlUkfdc.mjs';
import { defineComponent, ref, shallowRef, computed, watch, nextTick, mergeProps, unref, createVNode, resolveDynamicComponent, withCtx, openBlock, createBlock, Fragment, renderList, toDisplayString, useSSRContext } from 'vue';
import { ssrRenderAttrs, ssrIncludeBooleanAttr, ssrInterpolate, ssrRenderStyle, ssrRenderList, ssrRenderAttr, ssrRenderClass, ssrRenderComponent, ssrRenderVNode } from 'vue/server-renderer';
import 'nostics';
import 'nostics/formatters/ansi';
import '../_/nitro.mjs';
import 'node:http';
import 'node:https';
import 'node:events';
import 'node:buffer';
import 'node:fs';
import 'node:path';
import 'node:crypto';
import 'node:url';
import '../routes/renderer.mjs';
import 'unhead/server';
import 'unhead/legacy';
import 'unhead/plugins';
import 'vue-bundle-renderer/runtime';
import 'devalue';
import 'vue-router';
import 'unhead/utils';

var SpeakerNotes_vue_vue_type_script_setup_true_lang_default = /* @__PURE__ */ defineComponent({
  __name: "SpeakerNotes",
  __ssrInlineRender: true,
  props: { paragraphs: {} },
  emits: ["activate-hyperlink"],
  setup(__props, { emit: __emit }) {
    const emit = __emit;
    const spacing = (value) => value ? value.unit === "emu" ? `${value.value / 12700}pt` : `${value.value}em` : void 0;
    function paragraphStyle(paragraph) {
      var _a, _b;
      return {
        fontSize: ((_a = paragraph.runs[0]) == null ? void 0 : _a.fontSizePt) === void 0 ? void 0 : `${paragraph.runs[0].fontSizePt}pt`,
        color: (_b = paragraph.runs[0]) == null ? void 0 : _b.color,
        textAlign: paragraph.align,
        direction: paragraph.direction,
        marginLeft: paragraph.marginLeft === void 0 ? void 0 : `${paragraph.marginLeft / 12700}pt`,
        marginRight: paragraph.marginRight === void 0 ? void 0 : `${paragraph.marginRight / 12700}pt`,
        textIndent: paragraph.indent === void 0 ? void 0 : `${paragraph.indent / 12700}pt`,
        marginTop: spacing(paragraph.spaceBefore),
        marginBottom: spacing(paragraph.spaceAfter),
        lineHeight: spacing(paragraph.lineSpacing)
      };
    }
    function bulletStyle(paragraph) {
      var _a;
      const size = (_a = paragraph.bullet) == null ? void 0 : _a.size;
      return { fontSize: size ? size.unit === "ratio" ? `${size.value * 100}%` : `${size.value}pt` : void 0 };
    }
    function segmentFont(run, script) {
      return cssFontFamily((script === "eastAsia" ? run.fontFamilyEastAsia : script === "complexScript" ? run.fontFamilyComplexScript : run.fontFamily) || run.fontFamily);
    }
    function runStyle(run) {
      var _a, _b;
      const angle = (((_a = run.shadow) == null ? void 0 : _a.angle) || 0) * Math.PI / 180;
      const distance = (((_b = run.shadow) == null ? void 0 : _b.distance) || 0) / 12700;
      return {
        color: run.color,
        fontFamily: cssFontFamily(run.fontFamily),
        fontSize: run.fontSizePt === void 0 ? void 0 : `${run.fontSizePt}pt`,
        fontWeight: run.bold ? "bold" : "normal",
        fontStyle: run.italic ? "italic" : "normal",
        letterSpacing: run.letterSpacingPt === void 0 ? void 0 : `${run.letterSpacingPt}pt`,
        verticalAlign: run.baselineOffsetEm === void 0 ? void 0 : `${run.baselineOffsetEm}em`,
        textDecorationLine: [run.underline ? "underline" : "", run.lineThrough ? "line-through" : ""].filter(Boolean).join(" ") || "none",
        textDecorationStyle: run.underlineStyle === "wavy" ? "wavy" : run.underlineStyle || run.lineThroughStyle,
        WebkitTextStroke: run.outlineColor !== void 0 && run.outlineWidth !== void 0 ? `${run.outlineWidth / 12700}pt ${run.outlineColor}` : void 0,
        textShadow: run.shadow ? `${Math.cos(angle) * distance}pt ${Math.sin(angle) * distance}pt ${run.shadow.blur / 12700}pt ${run.shadow.color}` : void 0
      };
    }
    return (_ctx, _push, _parent, _attrs) => {
      _push(`<div${ssrRenderAttrs(mergeProps({
        class: "speaker-notes-rich",
        "data-presenter-notes": ""
      }, _attrs))} data-v-43fbb6b1><!--[-->`);
      ssrRenderList(unref(paragraphEntries)(__props.paragraphs), (entry, index) => {
        var _a;
        _push(`<p style="${ssrRenderStyle(paragraphStyle(entry.paragraph))}" data-v-43fbb6b1>`);
        if (((_a = entry.paragraph.bullet) == null ? void 0 : _a.kind) === "image" && entry.paragraph.bullet.url) _push(`<img class="note-picture-bullet"${ssrRenderAttr("src", entry.paragraph.bullet.url)} style="${ssrRenderStyle(bulletStyle(entry.paragraph))}" alt="" data-v-43fbb6b1>`);
        else if (entry.marker) _push(`<span class="note-bullet" style="${ssrRenderStyle(bulletStyle(entry.paragraph))}" data-v-43fbb6b1>${ssrInterpolate(entry.marker)}</span>`);
        else _push(`<!---->`);
        _push(`<!--[-->`);
        ssrRenderList(entry.paragraph.runs, (run, runIndex) => {
          ssrRenderVNode(_push, createVNode(resolveDynamicComponent(run.hyperlink ? "button" : "span"), {
            key: runIndex,
            lang: run.language,
            style: runStyle(run),
            type: run.hyperlink ? "button" : void 0,
            onClick: ($event) => run.hyperlink && emit("activate-hyperlink", run.hyperlink)
          }, {
            default: withCtx((_, _push2, _parent2, _scopeId) => {
              if (_push2) {
                _push2(`<!--[-->`);
                ssrRenderList(unref(scriptSegments)(run.text), (segment, segmentIndex) => {
                  _push2(`<span style="${ssrRenderStyle({ fontFamily: segmentFont(run, segment.script) })}" data-v-43fbb6b1${_scopeId}>${ssrInterpolate(segment.text)}</span>`);
                });
                _push2(`<!--]-->`);
              } else return [(openBlock(true), createBlock(Fragment, null, renderList(unref(scriptSegments)(run.text), (segment, segmentIndex) => {
                return openBlock(), createBlock("span", {
                  key: segmentIndex,
                  style: { fontFamily: segmentFont(run, segment.script) }
                }, toDisplayString(segment.text), 5);
              }), 128))];
            }),
            _: 2
          }), _parent);
        });
        _push(`<!--]-->${ssrInterpolate(index < __props.paragraphs.length - 1 ? "\n" : "")}</p>`);
      });
      _push(`<!--]--></div>`);
    };
  }
});
var _sfc_setup$1 = SpeakerNotes_vue_vue_type_script_setup_true_lang_default.setup;
SpeakerNotes_vue_vue_type_script_setup_true_lang_default.setup = (props, ctx) => {
  const ssrContext = useSSRContext();
  (ssrContext.modules || (ssrContext.modules = /* @__PURE__ */ new Set())).add("components/SpeakerNotes.vue");
  return _sfc_setup$1 ? _sfc_setup$1(props, ctx) : void 0;
};
var SpeakerNotes_default = /* @__PURE__ */ Object.assign(_plugin_vue_export_helper_default(SpeakerNotes_vue_vue_type_script_setup_true_lang_default, [["__scopeId", "data-v-43fbb6b1"]]), { __name: "SpeakerNotes" });
var thumbnailOverscan = 3;
var sorterPageSize = 48;
var index_vue_vue_type_script_setup_true_lang_default = /* @__PURE__ */ defineComponent({
  __name: "index",
  __ssrInlineRender: true,
  setup(__props) {
    ref(null);
    const deck = shallowRef(null);
    const deckGeneration = ref(0);
    const currentIndex = ref(0);
    const animationStep = ref(0);
    const errorMessage = ref("");
    const isLoading = ref(false);
    const isDragging = ref(false);
    const presenterView = ref(false);
    const sorterView = ref(false);
    const readingView = ref(false);
    const readingNotesOpen = ref(false);
    const presenterElapsedSeconds = ref(0);
    const projectorConnected = ref(false);
    const projectorOpen = ref(false);
    const projectorError = ref("");
    const projectorAcknowledgement = ref(null);
    const triggerPlayback = shallowRef([]);
    const transitionPlayback = shallowRef(null);
    ref(null);
    ref(null);
    ref(null);
    ref(null);
    ref(null);
    ref(null);
    const draggedSlideIndex = ref(null);
    const sorterPageStart = ref(0);
    const screenOverlay = ref(null);
    const thumbnailList = ref(null);
    const thumbnailStart = ref(0);
    const thumbnailEnd = ref(12);
    const thumbnailStep = ref(156);
    const thumbnailGap = ref(10);
    const thumbnailAxis = ref("vertical");
    const stageShell = ref(null);
    const stageSlideHost = ref(null);
    const transitionUnderlay = ref(null);
    const flashOverlay = ref(null);
    const presenterStageShell = ref(null);
    const presenterSlideHost = ref(null);
    const presenterTransitionUnderlay = ref(null);
    const presenterFlashOverlay = ref(null);
    const continuingMediaHost = ref(null);
    ref(null);
    ref(false);
    const slideNumberInput = ref("");
    const contextMenu = ref(null);
    const activeTool = ref("pointer");
    const inkBySlide = ref({});
    const draftStroke = ref(null);
    const laserPoint = ref(null);
    const zoom = ref(1);
    const zoomPan = ref({
      x: 0,
      y: 0
    });
    let slideNumberTimer;
    let slideAdvanceTimer;
    const continuingMedia = /* @__PURE__ */ new Map();
    const currentSlide = computed(() => {
      var _a;
      return ((_a = deck.value) == null ? void 0 : _a.slides[currentIndex.value]) || null;
    });
    const thumbnailSlides = computed(() => {
      var _a;
      return ((_a = deck.value) == null ? void 0 : _a.slides.slice(thumbnailStart.value, thumbnailEnd.value).map((slide, offset) => ({
        slide,
        index: thumbnailStart.value + offset
      }))) || [];
    });
    const thumbnailCount = computed(() => {
      var _a;
      return ((_a = deck.value) == null ? void 0 : _a.slides.length) || 0;
    });
    const thumbnailBeforeStyle = computed(() => thumbnailAxis.value === "vertical" ? { height: `${Math.max(0, thumbnailStart.value * thumbnailStep.value - thumbnailGap.value)}px` } : { width: `${Math.max(0, thumbnailStart.value * thumbnailStep.value - thumbnailGap.value)}px` });
    const thumbnailAfterStyle = computed(() => thumbnailAxis.value === "vertical" ? { height: `${Math.max(0, (thumbnailCount.value - thumbnailEnd.value) * thumbnailStep.value - thumbnailGap.value)}px` } : { width: `${Math.max(0, (thumbnailCount.value - thumbnailEnd.value) * thumbnailStep.value - thumbnailGap.value)}px` });
    const sorterSlides = computed(() => {
      var _a;
      return ((_a = deck.value) == null ? void 0 : _a.slides.slice(sorterPageStart.value, sorterPageStart.value + sorterPageSize).map((slide, offset) => ({
        slide,
        index: sorterPageStart.value + offset
      }))) || [];
    });
    const sorterPageEnd = computed(() => {
      var _a;
      return Math.min(sorterPageStart.value + sorterPageSize, ((_a = deck.value) == null ? void 0 : _a.slides.length) || 0);
    });
    const currentNumber = computed(() => deck.value ? currentIndex.value + 1 : 0);
    const animationCount = computed(() => {
      var _a, _b;
      return ((_b = (_a = currentSlide.value) == null ? void 0 : _a.animationSteps) == null ? void 0 : _b.length) || 0;
    });
    const progress = computed(() => deck.value ? currentNumber.value / deck.value.slides.length * 100 : 0);
    const currentInk = computed(() => currentSlide.value ? inkBySlide.value[currentSlide.value.id] || [] : []);
    const presenterTime = computed(() => `${String(Math.floor(presenterElapsedSeconds.value / 60)).padStart(2, "0")}:${String(presenterElapsedSeconds.value % 60).padStart(2, "0")}`);
    const randomTransitionEffects = [
      "fade",
      "wipe",
      "zoom",
      "circle",
      "diamond",
      "plus",
      "wedge",
      "newsflash",
      "honeycomb",
      "blinds",
      "checker",
      "comb",
      "dissolve",
      "wheel",
      "randomBar",
      "strips",
      "cover",
      "wheelReverse",
      "flash",
      "push"
    ];
    function createTransitionPlayback(slide) {
      if (!(slide == null ? void 0 : slide.transition)) return null;
      const source = slide.transition;
      const transition = source.effect === "random" ? {
        ...source,
        effect: randomTransitionEffects[Math.floor(Math.random() * randomTransitionEffects.length)]
      } : source;
      return {
        slideId: slide.id,
        startedAtEpoch: performance.timeOrigin + performance.now(),
        transition
      };
    }
    watch(currentIndex, () => {
      transitionPlayback.value = createTransitionPlayback(currentSlide.value);
      nextTick(ensureActiveThumbnailVisible);
    }, { flush: "sync" });
    useSlideTransition({
      slide: currentSlide,
      shell: stageShell,
      incomingHost: stageSlideHost,
      underlay: transitionUnderlay,
      flash: flashOverlay,
      playback: transitionPlayback
    });
    useSlideTransition({
      slide: currentSlide,
      shell: presenterStageShell,
      incomingHost: presenterSlideHost,
      underlay: presenterTransitionUnderlay,
      flash: presenterFlashOverlay,
      playback: transitionPlayback
    });
    function stopContinuingMedia(media, pause = true) {
      var _a;
      const state = continuingMedia.get(media);
      if (state == null ? void 0 : state.timer) clearTimeout(state.timer);
      if (state) media.removeEventListener("ended", state.onEnded);
      continuingMedia.delete(media);
      if (pause) media.pause();
      media.remove();
      (_a = state == null ? void 0 : state.frame) == null ? void 0 : _a.remove();
      if (!continuingMedia.size && continuingMediaHost.value) {
        continuingMediaHost.value.style.width = "0px";
        continuingMediaHost.value.style.height = "0px";
        continuingMediaHost.value.style.translate = "";
        continuingMediaHost.value.style.transform = "";
      }
    }
    function positionContinuingMediaHost() {
      const host = continuingMediaHost.value;
      const slide = stageSlideHost.value;
      if (!host || !slide || !continuingMedia.size) return;
      host.style.left = `${slide.offsetLeft}px`;
      host.style.top = `${slide.offsetTop}px`;
      host.style.width = `${slide.offsetWidth}px`;
      host.style.height = `${slide.offsetHeight}px`;
      host.style.translate = slide.style.translate;
      host.style.transform = slide.style.transform;
      host.style.transformOrigin = getComputedStyle(slide).transformOrigin;
    }
    function stopAllAudio() {
      for (const audio of (void 0).querySelectorAll("audio")) audio.pause();
      for (const media of [...continuingMedia.keys()]) stopContinuingMedia(media, false);
    }
    function continueMediaAcrossSlides(nextIndex, previousIndex) {
      var _a, _b;
      for (const [media, state] of continuingMedia) if (nextIndex < previousIndex || nextIndex >= state.endIndex || media.ended) stopContinuingMedia(media);
      for (const media of ((_a = stageShell.value) == null ? void 0 : _a.querySelectorAll("[data-media-id]")) || []) {
        if (continuingMedia.has(media)) continue;
        const slideCount = Number(media.dataset.mediaSlideCount);
        if (media.paused || media.ended || !Number.isInteger(slideCount) || slideCount <= 1 || nextIndex <= previousIndex || nextIndex >= previousIndex + slideCount) {
          media.pause();
          continue;
        }
        if (!(media instanceof HTMLAudioElement || media instanceof HTMLVideoElement) || !continuingMediaHost.value) {
          media.pause();
          continue;
        }
        const continuation = media.cloneNode(false);
        continuation.src = media.currentSrc || media.src;
        continuation.preload = "auto";
        continuation.volume = media.volume;
        continuation.muted = media.muted;
        continuation.playbackRate = media.playbackRate;
        continuation.loop = media.loop;
        continuation.addEventListener("click", (event) => event.stopPropagation());
        continuation.addEventListener("pointerdown", (event) => event.stopPropagation());
        continuation.dataset.mediaId = media.dataset.mediaId;
        if (media.dataset.mediaActionEndAt) continuation.dataset.mediaActionEndAt = media.dataset.mediaActionEndAt;
        const startAt = media.currentTime;
        const onEnded = () => stopContinuingMedia(continuation, false);
        let frame;
        if (continuation instanceof HTMLVideoElement) {
          const sourceFrame = media.closest(".slide-element");
          const slideFrame = (_b = stageSlideHost.value) == null ? void 0 : _b.querySelector(".slide-frame");
          if (!sourceFrame || !slideFrame || !slideFrame.offsetWidth || !slideFrame.offsetHeight) {
            media.pause();
            continue;
          }
          const sourceStyle = getComputedStyle(sourceFrame);
          frame = (void 0).createElement("div");
          frame.className = "continuing-video-frame";
          Object.assign(frame.style, {
            left: `${sourceFrame.offsetLeft / slideFrame.offsetWidth * 100}%`,
            top: `${sourceFrame.offsetTop / slideFrame.offsetHeight * 100}%`,
            width: `${sourceFrame.offsetWidth / slideFrame.offsetWidth * 100}%`,
            height: `${sourceFrame.offsetHeight / slideFrame.offsetHeight * 100}%`,
            transform: sourceStyle.transform,
            transformOrigin: sourceStyle.transformOrigin,
            opacity: sourceStyle.opacity,
            clipPath: sourceStyle.clipPath
          });
          frame.append(continuation);
        }
        const state = {
          endIndex: previousIndex + slideCount,
          onEnded
        };
        if (frame) state.frame = frame;
        const endAt = Number(continuation.dataset.mediaActionEndAt);
        if (Number.isFinite(endAt)) state.timer = setTimeout(() => stopContinuingMedia(continuation), Math.max(0, endAt - performance.now()));
        continuation.addEventListener("ended", onEnded, { once: true });
        continuingMedia.set(continuation, state);
        if (frame) continuingMediaHost.value.append(frame);
        else continuingMediaHost.value.append(continuation);
        positionContinuingMediaHost();
        const start = () => {
          if (!continuingMedia.has(continuation)) return;
          try {
            continuation.currentTime = startAt;
          } catch {
          }
          continuation.play().then(() => media.pause()).catch(() => stopContinuingMedia(continuation, false));
        };
        if (continuation.readyState >= 1) start();
        else continuation.addEventListener("loadedmetadata", start, { once: true });
      }
      nextTick(positionContinuingMediaHost);
    }
    function sendProjectorDocument() {
      if (!deck.value) return;
      deck.value.name, deck.value.width, deck.value.height, deck.value.slides.length, deck.value.embeddedFonts, deck.value.warnings.filter((warning) => /embedded font/i.test(warning));
    }
    function sendProjectorSlide() {
      if (!currentSlide.value) return;
      currentSlide.value, currentNumber.value, animationStep.value, screenOverlay.value, triggerPlayback.value, transitionPlayback.value, currentInk.value.map((stroke) => ({ ...stroke })), draftStroke.value && { ...draftStroke.value }, activeTool.value === "laser" && laserPoint.value && { ...laserPoint.value };
    }
    function sendProjectorState() {
      animationStep.value, screenOverlay.value, triggerPlayback.value;
    }
    function sendProjectorLaserPoint(point) {
      if (currentSlide.value) currentSlide.value.id;
    }
    watch(deck, () => {
      projectorAcknowledgement.value = null;
      projectorConnected.value = false;
      sendProjectorDocument();
      nextTick(() => {
        if (thumbnailList.value) {
          thumbnailList.value.scrollTop = 0;
          thumbnailList.value.scrollLeft = 0;
        }
        thumbnailStart.value = 0;
        thumbnailEnd.value = Math.min(thumbnailCount.value, 12);
        updateThumbnailWindow();
      });
    }, { flush: "post" });
    watch([deck, currentIndex], () => {
      draftStroke.value = null;
      laserPoint.value = null;
      triggerPlayback.value = [];
      sendProjectorSlide();
    }, { flush: "post" });
    watch([
      animationStep,
      screenOverlay,
      triggerPlayback
    ], sendProjectorState, { flush: "post" });
    watch(activeTool, (tool) => {
      if (tool !== "laser") {
        laserPoint.value = null;
        sendProjectorLaserPoint();
      }
    });
    function clearSlideNumberInput() {
      slideNumberInput.value = "";
      clearTimeout(slideNumberTimer);
      slideNumberTimer = void 0;
    }
    function nextSlide() {
      var _a, _b;
      clearSlideNumberInput();
      if (animationStep.value < animationCount.value) {
        if (((_b = (_a = currentSlide.value) == null ? void 0 : _a.animationSequence) == null ? void 0 : _b.advancesOnNext) === false) return;
        animationStep.value++;
      } else if (deck.value && currentIndex.value < deck.value.slides.length - 1) {
        currentIndex.value++;
        animationStep.value = 0;
      }
    }
    function goToSlide(index) {
      clearSlideNumberInput();
      currentIndex.value = index;
      animationStep.value = 0;
    }
    function activateHyperlink(hyperlink) {
      if (!deck.value) return;
      if (hyperlink.kind === "external") {
        try {
          const url = new URL(hyperlink.url);
          if ([
            "http:",
            "https:",
            "mailto:"
          ].includes(url.protocol) && !url.username && !url.password) (void 0).open(url.href, "_blank", "noopener,noreferrer");
        } catch {
        }
        return;
      }
      let targetIndex;
      if (hyperlink.kind === "slide") targetIndex = deck.value.slides.findIndex((slide) => slide.packagePath === hyperlink.targetPath);
      else targetIndex = hyperlink.jump === "firstslide" ? 0 : hyperlink.jump === "lastslide" ? deck.value.slides.length - 1 : hyperlink.jump === "nextslide" ? currentIndex.value + 1 : hyperlink.jump === "previousslide" ? currentIndex.value - 1 : (hyperlink.slideNumber || 0) - 1;
      if (targetIndex >= 0 && targetIndex < deck.value.slides.length) goToSlide(targetIndex);
    }
    function updateThumbnailWindow() {
      var _a;
      const list = thumbnailList.value;
      const total = thumbnailCount.value;
      if (!list || !deck.value) {
        thumbnailStart.value = 0;
        thumbnailEnd.value = 0;
        return;
      }
      const axis = (void 0).matchMedia("(max-width: 860px)").matches ? "horizontal" : "vertical";
      const style = getComputedStyle(list);
      const gap = Number.parseFloat(axis === "horizontal" ? style.columnGap : style.rowGap) || 0;
      const size = (_a = list.querySelector(".thumb-button")) == null ? void 0 : _a.getBoundingClientRect()[axis === "horizontal" ? "width" : "height"];
      const step = size ? size + gap : thumbnailStep.value;
      thumbnailAxis.value = axis;
      thumbnailGap.value = gap;
      if (step > gap) thumbnailStep.value = step;
      const offset = axis === "horizontal" ? list.scrollLeft : list.scrollTop;
      const viewport = axis === "horizontal" ? list.clientWidth : list.clientHeight;
      const start = Math.max(0, Math.floor(offset / thumbnailStep.value) - thumbnailOverscan);
      const end = Math.min(total, Math.max(start + 1, Math.ceil((offset + viewport) / thumbnailStep.value) + thumbnailOverscan));
      if (start !== thumbnailStart.value) thumbnailStart.value = start;
      if (end !== thumbnailEnd.value) thumbnailEnd.value = end;
    }
    function ensureActiveThumbnailVisible() {
      const list = thumbnailList.value;
      const total = thumbnailCount.value;
      if (!list || !total) return;
      const hadFocus = list.contains((void 0).activeElement);
      const horizontal = (void 0).matchMedia("(max-width: 860px)").matches;
      const style = getComputedStyle(list);
      const paddingStart = Number.parseFloat(horizontal ? style.paddingLeft : style.paddingTop) || 0;
      const paddingEnd = Number.parseFloat(horizontal ? style.paddingRight : style.paddingBottom) || 0;
      const viewport = horizontal ? list.clientWidth : list.clientHeight;
      const offset = horizontal ? list.scrollLeft : list.scrollTop;
      const index = Math.min(total - 1, Math.max(0, currentIndex.value));
      const itemStart = paddingStart + index * thumbnailStep.value;
      const itemEnd = itemStart + thumbnailStep.value - thumbnailGap.value;
      let nextOffset = offset;
      if (itemStart < offset + paddingStart) nextOffset = index * thumbnailStep.value;
      else if (itemEnd > offset + viewport - paddingEnd) nextOffset = itemEnd - viewport + paddingEnd;
      if (horizontal) list.scrollLeft = Math.max(0, nextOffset);
      else list.scrollTop = Math.max(0, nextOffset);
      updateThumbnailWindow();
      if (hadFocus) nextTick(() => {
        var _a;
        return (_a = list.querySelector(`[data-thumbnail-index="${index}"]`)) == null ? void 0 : _a.focus({ preventScroll: true });
      });
    }
    function setZoom(value) {
      zoom.value = Math.min(3, Math.max(0.5, Math.round(value * 100) / 100));
      setZoomPan(zoomPan.value.x, zoomPan.value.y);
    }
    function setZoomPan(x, y) {
      const host = stageSlideHost.value;
      if (zoom.value <= 1 || !host) {
        zoomPan.value = {
          x: 0,
          y: 0
        };
        return;
      }
      const maxX = host.offsetWidth * (zoom.value - 1) / 2;
      const maxY = host.offsetHeight * (zoom.value - 1) / 2;
      zoomPan.value = {
        x: Math.max(-maxX, Math.min(maxX, x)),
        y: Math.max(-maxY, Math.min(maxY, y))
      };
    }
    function scheduleSlideAdvance() {
      var _a;
      if (slideAdvanceTimer) clearTimeout(slideAdvanceTimer);
      slideAdvanceTimer = void 0;
      const slide = currentSlide.value;
      const delayMs = slide == null ? void 0 : slide.advanceAfterMs;
      if (!slide || delayMs === void 0 || currentIndex.value >= (((_a = deck.value) == null ? void 0 : _a.slides.length) || 0) - 1) return;
      const slideIndex = currentIndex.value;
      const dueAt = Date.now() + delayMs;
      const advance = () => {
        const remaining = dueAt - Date.now();
        if (remaining > 0) {
          slideAdvanceTimer = setTimeout(advance, Math.min(remaining, 2147483647));
          return;
        }
        if (currentIndex.value !== slideIndex || currentSlide.value !== slide) return;
        animationStep.value = animationCount.value;
        nextSlide();
      };
      slideAdvanceTimer = setTimeout(advance, Math.min(delayMs, 2147483647));
    }
    watch(currentIndex, (nextIndex, previousIndex) => {
      if (nextIndex === previousIndex) return;
      continueMediaAcrossSlides(nextIndex, previousIndex);
      draftStroke.value = null;
      laserPoint.value = null;
      setZoom(1);
      nextTick(positionContinuingMediaHost);
    }, { flush: "sync" });
    watch([
      zoom,
      () => zoomPan.value.x,
      () => zoomPan.value.y
    ], () => {
      nextTick(positionContinuingMediaHost);
    }, { flush: "post" });
    watch([() => deck.value, () => currentIndex.value], scheduleSlideAdvance, { immediate: true });
    return (_ctx, _push, _parent, _attrs) => {
      var _a, _b, _c, _d, _e, _f, _g, _h, _i, _j, _k, _l, _m, _n, _o, _p, _q, _r, _s, _t, _u, _v, _w, _x, _y;
      const _component_SlideStage = SlideStage_default;
      const _component_SpeakerNotes = SpeakerNotes_default;
      _push(`<main${ssrRenderAttrs(mergeProps({ class: ["app-shell", { "is-reading-view": unref(readingView) }] }, _attrs))}><header class="topbar"><div class="brand" aria-label="Deckline PPTX Player"><span class="brand-mark">D</span><span class="brand-name">deckline</span><span class="brand-label">PPTX PLAYER</span></div><div class="topbar-status"><span class="status-dot"></span>LOCAL SESSION \xB7 FILES STAY ON DEVICE</div></header><div class="workspace"${ssrIncludeBooleanAttr(unref(presenterView) || unref(sorterView)) ? " inert" : ""}><aside class="deck-sidebar" aria-label="Slide list"><div class="sidebar-head"><span class="sidebar-title">Slides</span><span class="slide-count">${ssrInterpolate(unref(deck) ? String(unref(deck).slides.length).padStart(2, "0") : "00")}</span></div>`);
      if (unref(deck)) {
        _push(`<div class="thumb-list">`);
        if (unref(thumbnailStart) > 0) _push(`<div class="thumb-virtual-spacer" style="${ssrRenderStyle(unref(thumbnailBeforeStyle))}" aria-hidden="true"></div>`);
        else _push(`<!---->`);
        _push(`<!--[-->`);
        ssrRenderList(unref(thumbnailSlides), ({ slide, index }) => {
          _push(`<button${ssrRenderAttr("data-thumbnail-index", index)} class="${ssrRenderClass([{ "is-active": unref(currentIndex) === index }, "thumb-button"])}"${ssrRenderAttr("aria-current", unref(currentIndex) === index ? "page" : void 0)}${ssrRenderAttr("aria-label", `Go to slide ${index + 1}`)}><div class="thumb-preview">`);
          _push(ssrRenderComponent(_component_SlideStage, {
            slide,
            width: unref(deck).width,
            height: unref(deck).height,
            number: index + 1,
            thumbnail: ""
          }, null, _parent));
          _push(`</div><div class="thumb-meta"><span>SLIDE</span><span>${ssrInterpolate(String(index + 1).padStart(2, "0"))}</span></div></button>`);
        });
        _push(`<!--]-->`);
        if (unref(thumbnailEnd) < unref(thumbnailCount)) _push(`<div class="thumb-virtual-spacer" style="${ssrRenderStyle(unref(thumbnailAfterStyle))}" aria-hidden="true"></div>`);
        else _push(`<!---->`);
        _push(`</div>`);
      } else _push(`<div class="side-empty">Your slide thumbnails will appear here after you open a presentation.</div>`);
      _push(`<div class="sidebar-footer">POWERED BY OOXML \xB7 BROWSER ONLY</div></aside><section class="player-column" aria-label="Presentation player"><div class="deck-toolbar"><span class="file-icon">\u25A4</span><div class="deck-title-block"><div class="deck-title">${ssrInterpolate(((_a = unref(deck)) == null ? void 0 : _a.name) || "No presentation open")}</div><div class="deck-subtitle">${ssrInterpolate(unref(deck) ? `${unref(deck).slides.length} SLIDES \xB7 READY TO PRESENT` : "OPEN A POWERPOINT FILE TO BEGIN")}</div></div><div class="tool-separator"></div><button class="icon-button"${ssrIncludeBooleanAttr(!unref(deck) || unref(currentIndex) === 0 && unref(animationStep) === 0) ? " disabled" : ""} aria-label="Previous animation step or slide" title="Previous animation step or slide">\u2039</button><button class="icon-button"${ssrIncludeBooleanAttr(!unref(deck) || unref(currentIndex) === (((_b = unref(deck)) == null ? void 0 : _b.slides.length) || 0) - 1 && unref(animationStep) >= unref(animationCount)) ? " disabled" : ""} aria-label="Next animation step or slide" title="Next animation step or slide">\u203A</button><button class="text-button"${ssrIncludeBooleanAttr(!unref(deck)) ? " disabled" : ""} title="Fullscreen (F)">\u26F6 <span>Present</span></button><button class="tool-mode-button" data-tool="sorter"${ssrIncludeBooleanAttr(!unref(deck)) ? " disabled" : ""} aria-label="Open slide sorter" title="Slide sorter (G)">sort</button><button class="tool-mode-button" data-tool="speaker"${ssrIncludeBooleanAttr(!unref(deck)) ? " disabled" : ""} aria-label="Open speaker view" title="Speaker view (S)">speaker</button><button class="tool-mode-button" data-tool="reading"${ssrIncludeBooleanAttr(!unref(deck)) ? " disabled" : ""} aria-label="Open reading view" title="Reading view">read</button><!--[-->`);
      ssrRenderList([
        "pointer",
        "pen",
        "highlighter",
        "laser"
      ], (tool) => {
        _push(`<button class="${ssrRenderClass([{ "is-active": unref(activeTool) === tool }, "tool-mode-button"])}"${ssrRenderAttr("data-tool", tool)}${ssrIncludeBooleanAttr(!unref(deck)) ? " disabled" : ""}>${ssrInterpolate(tool)}</button>`);
      });
      _push(`<!--]--><button class="tool-mode-button"${ssrIncludeBooleanAttr(!unref(deck) || !unref(currentInk).length) ? " disabled" : ""} data-tool="clear">clear</button><div class="tool-separator"></div><button class="primary-button"${ssrIncludeBooleanAttr(unref(isLoading)) ? " disabled" : ""}>${ssrInterpolate(unref(isLoading) ? "Reading\u2026" : unref(deck) ? "Open another" : "Open PPTX")}</button><input class="visually-hidden" type="file" tabindex="-1" aria-hidden="true" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"></div><div class="stage-shell slide-transition-shell"><div class="stage-grid"></div><div class="${ssrRenderClass([{
        "is-pulling": ((_d = (_c = unref(currentSlide)) == null ? void 0 : _c.transition) == null ? void 0 : _d.effect) === "pull",
        "is-splitting-in": ((_f = (_e = unref(currentSlide)) == null ? void 0 : _e.transition) == null ? void 0 : _f.effect) === "split" && unref(currentSlide).transition.direction === "in"
      }, "slide-transition-underlay"])}" aria-hidden="true"></div><div class="continuing-media-host"></div>`);
      if (unref(isDragging)) _push(`<div class="upload-overlay">DROP YOUR PRESENTATION</div>`);
      else _push(`<!---->`);
      if (unref(currentSlide) && unref(deck)) {
        _push(`<div class="slide-host"${ssrRenderAttr("data-slide-number", unref(currentNumber))}${ssrRenderAttr("data-transition-effect", (_g = unref(currentSlide).transition) == null ? void 0 : _g.effect)}${ssrRenderAttr("data-transition-preset", (_h = unref(currentSlide).transition) == null ? void 0 : _h.presetName)} style="${ssrRenderStyle({
          "--slide-ratio": String(unref(deck).width / unref(deck).height),
          translate: unref(zoomPan).x + "px " + unref(zoomPan).y + "px",
          transform: "scale(" + unref(zoom) + ")"
        })}">`);
        _push(ssrRenderComponent(_component_SlideStage, {
          key: unref(deckGeneration) + "-" + unref(currentSlide).id,
          slide: unref(currentSlide),
          width: unref(deck).width,
          height: unref(deck).height,
          number: unref(currentNumber),
          "animation-step": unref(animationStep),
          onTriggerPlayback: ($event) => triggerPlayback.value = $event,
          onActivateHyperlink: activateHyperlink,
          onStopAllAudio: stopAllAudio
        }, null, _parent));
        _push(`<svg class="${ssrRenderClass([{ "is-active": unref(activeTool) !== "pointer" }, "ink-layer"])}"${ssrRenderAttr("viewBox", `0 0 ${unref(deck).width} ${unref(deck).height}`)} preserveAspectRatio="none" aria-label="Presentation annotation layer"><!--[-->`);
        ssrRenderList(unref(currentInk), (stroke, index) => {
          _push(`<polyline${ssrRenderAttr("points", stroke.points)} fill="none"${ssrRenderAttr("stroke", stroke.tool === "pen" ? "#ef4444" : "#fde047")}${ssrRenderAttr("stroke-width", stroke.tool === "pen" ? 3 : 16)}${ssrRenderAttr("opacity", stroke.tool === "pen" ? 1 : 0.4)} stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke"></polyline>`);
        });
        _push(`<!--]-->`);
        if (unref(draftStroke)) _push(`<polyline${ssrRenderAttr("points", unref(draftStroke).points)} fill="none"${ssrRenderAttr("stroke", unref(draftStroke).tool === "pen" ? "#ef4444" : "#fde047")}${ssrRenderAttr("stroke-width", unref(draftStroke).tool === "pen" ? 3 : 16)}${ssrRenderAttr("opacity", unref(draftStroke).tool === "pen" ? 1 : 0.4)} stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke"></polyline>`);
        else _push(`<!---->`);
        if (unref(laserPoint)) _push(`<circle${ssrRenderAttr("cx", unref(laserPoint).x)}${ssrRenderAttr("cy", unref(laserPoint).y)}${ssrRenderAttr("r", unref(deck).width / 120)} fill="#ef4444" opacity=".9"></circle>`);
        else _push(`<!---->`);
        _push(`</svg><div class="flash-transition-overlay" aria-hidden="true"></div></div>`);
      } else if (!unref(errorMessage)) _push(`<div class="empty-state"><div class="empty-art"><span class="empty-art-back"></span><span class="empty-art-front"></span></div><div class="empty-kicker">A quieter way to present</div><h1>Your deck, in the room.</h1><p>Open a PowerPoint file and step through your slides. The presentation is read locally in your browser.</p><button class="primary-button"${ssrIncludeBooleanAttr(unref(isLoading)) ? " disabled" : ""}>${ssrInterpolate(unref(isLoading) ? "Reading presentation\u2026" : "Choose a .pptx file")}</button><div class="drop-hint">OR DROP A FILE ANYWHERE IN THIS WINDOW</div></div>`);
      else _push(`<!---->`);
      if (unref(screenOverlay)) _push(`<div class="${ssrRenderClass([`is-${unref(screenOverlay)}`, "screen-overlay"])}"></div>`);
      else _push(`<!---->`);
      if (unref(slideNumberInput)) _push(`<div class="slide-number-input" aria-live="polite">${ssrInterpolate(unref(slideNumberInput))}</div>`);
      else _push(`<!---->`);
      if (unref(contextMenu)) _push(`<div class="presentation-menu" style="${ssrRenderStyle({
        left: `${unref(contextMenu).x}px`,
        top: `${unref(contextMenu).y}px`
      })}" role="menu"><button role="menuitem"${ssrIncludeBooleanAttr(unref(currentIndex) === 0 && unref(animationStep) === 0) ? " disabled" : ""}>Previous</button><button role="menuitem"${ssrIncludeBooleanAttr(unref(currentIndex) === (((_i = unref(deck)) == null ? void 0 : _i.slides.length) || 0) - 1 && unref(animationStep) >= unref(animationCount)) ? " disabled" : ""}>Next</button><button role="menuitem">Black screen</button><button role="menuitem">White screen</button></div>`);
      else _push(`<!---->`);
      if (unref(errorMessage)) _push(`<div class="error-card" role="alert"><strong>Couldn\u2019t open this presentation</strong>${ssrInterpolate(unref(errorMessage))}</div>`);
      else _push(`<!---->`);
      if ((_j = unref(deck)) == null ? void 0 : _j.warnings.length) {
        _push(`<div class="warning-strip" aria-label="Rendering limitations"><!--[-->`);
        ssrRenderList(unref(deck).warnings, (warning) => {
          _push(`<span class="warning-pill">${ssrInterpolate(warning)}</span>`);
        });
        _push(`<!--]--></div>`);
      } else _push(`<!---->`);
      if (unref(deck)) _push(`<div class="progress-track" aria-hidden="true"><div class="progress-value" style="${ssrRenderStyle({ width: `${unref(progress)}%` })}"></div></div>`);
      else _push(`<!---->`);
      _push(`</div><div class="player-controls"><button class="icon-button"${ssrIncludeBooleanAttr(!unref(deck) || unref(currentIndex) === 0 && unref(animationStep) === 0) ? " disabled" : ""} aria-label="Previous animation step or slide">\u2190</button><div class="slide-position"><strong>${ssrInterpolate(String(unref(currentNumber)).padStart(2, "0"))}</strong> / ${ssrInterpolate(String(((_k = unref(deck)) == null ? void 0 : _k.slides.length) || 0).padStart(2, "0"))}</div><button class="icon-button"${ssrIncludeBooleanAttr(!unref(deck) || unref(currentIndex) === (((_l = unref(deck)) == null ? void 0 : _l.slides.length) || 0) - 1 && unref(animationStep) >= unref(animationCount)) ? " disabled" : ""} aria-label="Next animation step or slide">\u2192</button><button class="zoom-button"${ssrIncludeBooleanAttr(!unref(deck) || unref(zoom) <= 0.5) ? " disabled" : ""} aria-label="Zoom out">\u2212</button><span class="zoom-value">${ssrInterpolate(Math.round(unref(zoom) * 100))}%</span><button class="zoom-button"${ssrIncludeBooleanAttr(!unref(deck) || unref(zoom) >= 3) ? " disabled" : ""} aria-label="Zoom in">+</button><span class="shortcut-hint">ARROWS \xB7 F5 START \xB7 SHIFT+F5 CURRENT \xB7 F FULLSCREEN</span>`);
      if (unref(readingView)) _push(`<button class="reading-control-button" data-reading-notes${ssrRenderAttr("aria-expanded", unref(readingNotesOpen))}${ssrRenderAttr("aria-controls", unref(readingNotesOpen) ? "reading-notes" : void 0)}>${ssrInterpolate(unref(readingNotesOpen) ? "Hide notes" : "Notes")}</button>`);
      else _push(`<!---->`);
      if (unref(readingView)) _push(`<button class="reading-control-button" data-reading-exit>Back to player</button>`);
      else _push(`<!---->`);
      _push(`</div>`);
      if (unref(readingView) && unref(readingNotesOpen) && unref(currentSlide)) {
        _push(`<aside id="reading-notes" class="reading-notes-panel" data-reading-notes-panel aria-label="Speaker notes"><header><strong>Speaker notes</strong><button aria-label="Hide speaker notes">\xD7</button></header>`);
        if ((_m = unref(currentSlide).speakerNoteParagraphs) == null ? void 0 : _m.length) _push(ssrRenderComponent(_component_SpeakerNotes, {
          paragraphs: unref(currentSlide).speakerNoteParagraphs,
          onActivateHyperlink: activateHyperlink
        }, null, _parent));
        else if (unref(currentSlide).speakerNotes) _push(`<p class="reading-notes-plain">${ssrInterpolate(unref(currentSlide).speakerNotes)}</p>`);
        else _push(`<p class="reading-notes-plain">No speaker notes on this slide.</p>`);
        _push(`</aside>`);
      } else _push(`<!---->`);
      _push(`</section></div>`);
      if (unref(presenterView) && unref(deck) && unref(currentSlide)) {
        _push(`<div class="presenter-view" data-presenter-view${ssrRenderAttr("data-current-slide", unref(currentNumber))}${ssrRenderAttr("data-animation-step", unref(animationStep))} role="dialog" aria-modal="true" aria-label="Speaker view"><header class="presenter-header"><div><strong>${ssrInterpolate(unref(deck).name)}</strong><span> SPEAKER VIEW</span></div><div class="presenter-header-actions"><span class="presenter-clock" aria-label="Elapsed time">${ssrInterpolate(unref(presenterTime))}</span><span class="presenter-display-status"${ssrRenderAttr("data-projector-connected", unref(projectorConnected))}${ssrRenderAttr("data-projector-number", (_o = (_n = unref(projectorAcknowledgement)) == null ? void 0 : _n.number) != null ? _o : 0)}${ssrRenderAttr("data-projector-step", (_q = (_p = unref(projectorAcknowledgement)) == null ? void 0 : _p.animationStep) != null ? _q : -1)}${ssrRenderAttr("data-projector-blank", ((_r = unref(projectorAcknowledgement)) == null ? void 0 : _r.blank) || "")}${ssrRenderAttr("data-projector-trigger-count", (_t = (_s = unref(projectorAcknowledgement)) == null ? void 0 : _s.triggerCount) != null ? _t : -1)}>${ssrInterpolate(unref(projectorConnected) ? "SECOND DISPLAY CONNECTED" : "SECOND DISPLAY CLOSED")}</span><button class="text-button" data-open-projector>${ssrInterpolate(unref(projectorOpen) ? "Focus display" : "Open second display")}</button>`);
        if (unref(projectorError)) _push(`<span class="projector-error" data-projector-error role="alert">${ssrInterpolate(unref(projectorError))}</span>`);
        else _push(`<!---->`);
        if (unref(projectorOpen)) _push(`<button class="text-button" data-close-projector>Close display</button>`);
        else _push(`<!---->`);
        _push(`<button class="text-button" data-presenter-exit>Exit <kbd>Esc</kbd></button></div></header><div class="presenter-grid"><section class="presenter-current-panel" aria-label="Current slide"><div class="presenter-panel-title"><span>Current</span><span>${ssrInterpolate(unref(currentNumber))} / ${ssrInterpolate(unref(deck).slides.length)}</span></div><div class="presenter-current-frame"><div class="slide-transition-shell presenter-transition-shell"><div class="${ssrRenderClass([{
          "is-pulling": ((_v = (_u = unref(currentSlide)) == null ? void 0 : _u.transition) == null ? void 0 : _v.effect) === "pull",
          "is-splitting-in": ((_x = (_w = unref(currentSlide)) == null ? void 0 : _w.transition) == null ? void 0 : _x.effect) === "split" && unref(currentSlide).transition.direction === "in"
        }, "slide-transition-underlay"])}" aria-hidden="true"></div><div class="slide-host"${ssrRenderAttr("data-slide-number", unref(currentNumber))} style="${ssrRenderStyle({ "--slide-ratio": String(unref(deck).width / unref(deck).height) })}">`);
        _push(ssrRenderComponent(_component_SlideStage, {
          class: "presenter-large-slide",
          style: { "--slide-ratio": String(unref(deck).width / unref(deck).height) },
          slide: unref(currentSlide),
          width: unref(deck).width,
          height: unref(deck).height,
          number: unref(currentNumber),
          "animation-step": unref(animationStep),
          "trigger-playback": unref(triggerPlayback),
          thumbnail: "",
          "playback-preview": ""
        }, null, _parent));
        _push(`</div><div class="flash-transition-overlay" aria-hidden="true"></div></div></div></section><aside class="presenter-side-panel"><section class="presenter-next-panel" aria-label="Next slide"><div class="presenter-panel-title"><span>Next</span><span>${ssrInterpolate(unref(currentIndex) < unref(deck).slides.length - 1 ? `${unref(currentIndex) + 2} / ${unref(deck).slides.length}` : "End of deck")}</span></div>`);
        if (unref(deck).slides[unref(currentIndex) + 1]) {
          _push(`<div class="presenter-next-frame">`);
          _push(ssrRenderComponent(_component_SlideStage, {
            class: "presenter-next-slide",
            slide: unref(deck).slides[unref(currentIndex) + 1],
            width: unref(deck).width,
            height: unref(deck).height,
            number: unref(currentIndex) + 2,
            thumbnail: ""
          }, null, _parent));
          _push(`</div>`);
        } else _push(`<div class="presenter-end-card">This is the final slide.</div>`);
        _push(`</section><section class="presenter-notes-panel" aria-label="Speaker notes"><div class="presenter-panel-title"><span>Speaker notes</span></div>`);
        if ((_y = unref(currentSlide).speakerNoteParagraphs) == null ? void 0 : _y.length) _push(ssrRenderComponent(_component_SpeakerNotes, {
          paragraphs: unref(currentSlide).speakerNoteParagraphs,
          onActivateHyperlink: activateHyperlink
        }, null, _parent));
        else if (unref(currentSlide).speakerNotes) _push(`<p data-presenter-notes>${ssrInterpolate(unref(currentSlide).speakerNotes)}</p>`);
        else _push(`<p class="presenter-no-notes" data-presenter-notes>No speaker notes on this slide.</p>`);
        _push(`</section></aside></div><footer class="presenter-footer"><button class="text-button"${ssrIncludeBooleanAttr(unref(currentIndex) === 0 && unref(animationStep) === 0) ? " disabled" : ""}>\u2190 Previous</button><span>Arrow keys navigate \xB7 S toggles speaker view \xB7 ${ssrInterpolate(unref(presenterTime))}</span><button class="text-button"${ssrIncludeBooleanAttr(unref(currentIndex) === unref(deck).slides.length - 1 && unref(animationStep) >= unref(animationCount)) ? " disabled" : ""}>Next \u2192</button></footer></div>`);
      } else _push(`<!---->`);
      if (unref(sorterView) && unref(deck)) {
        _push(`<div class="sorter-view" data-sorter-view role="dialog" aria-modal="true" aria-label="Slide sorter"><header class="sorter-header"><div><strong>Slide sorter</strong><span>${ssrInterpolate(unref(sorterPageStart) + 1)}\u2013${ssrInterpolate(unref(sorterPageEnd))} / ${ssrInterpolate(unref(deck).slides.length)}</span></div><button class="text-button" data-sorter-exit>Close <kbd>Esc</kbd></button></header><div class="sorter-grid" role="list"><!--[-->`);
        ssrRenderList(unref(sorterSlides), ({ slide, index }) => {
          _push(`<article class="${ssrRenderClass([{
            "is-active": unref(currentIndex) === index,
            "is-dragging": unref(draggedSlideIndex) === index
          }, "sorter-card"])}"${ssrRenderAttr("data-slide-index", index)} role="listitem" draggable="true"><button class="sorter-slide-button"${ssrRenderAttr("aria-label", `Go to slide ${index + 1}: ${slide.name}`)}><div class="sorter-preview" style="${ssrRenderStyle({ aspectRatio: `${unref(deck).width} / ${unref(deck).height}` })}">`);
          _push(ssrRenderComponent(_component_SlideStage, {
            slide,
            width: unref(deck).width,
            height: unref(deck).height,
            number: index + 1,
            thumbnail: ""
          }, null, _parent));
          _push(`</div><div class="sorter-meta"><span>${ssrInterpolate(String(index + 1).padStart(2, "0"))}</span><span>${ssrInterpolate(slide.name)}</span></div></button></article>`);
        });
        _push(`<!--]--></div><footer class="sorter-footer"><button class="text-button"${ssrIncludeBooleanAttr(unref(sorterPageStart) === 0) ? " disabled" : ""}>\u2190 Previous group</button><span>Drag to reorder in this group \xB7 Select a slide, then use Alt + arrow keys to move it</span><button class="text-button"${ssrIncludeBooleanAttr(unref(sorterPageEnd) >= unref(deck).slides.length) ? " disabled" : ""}>Next group \u2192</button></footer></div>`);
      } else _push(`<!---->`);
      _push(`</main>`);
    };
  }
});
var _sfc_setup = index_vue_vue_type_script_setup_true_lang_default.setup;
index_vue_vue_type_script_setup_true_lang_default.setup = (props, ctx) => {
  const ssrContext = useSSRContext();
  (ssrContext.modules || (ssrContext.modules = /* @__PURE__ */ new Set())).add("pages/index.vue");
  return _sfc_setup ? _sfc_setup(props, ctx) : void 0;
};
var pages_default = index_vue_vue_type_script_setup_true_lang_default;

export { pages_default as default };
//# sourceMappingURL=pages-C88-SEg7.mjs.map
