import { u as useRoute$1 } from '../virtual/entry.mjs';
import { u as useSlideTransition, S as SlideStage_default } from './useSlideTransition-DcddA0qZ.mjs';
import { defineComponent, shallowRef, ref, computed, watch, mergeProps, unref, useSSRContext } from 'vue';
import { ssrRenderAttrs, ssrRenderClass, ssrRenderAttr, ssrRenderStyle, ssrRenderComponent, ssrRenderList, ssrInterpolate } from 'vue/server-renderer';
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

//#region app/pages/projector.vue?vue&type=script&setup=true&lang.ts
var projector_vue_vue_type_script_setup_true_lang_default = /*@__PURE__*/ defineComponent({
	__name: "projector",
	__ssrInlineRender: true,
	setup(__props) {
		useRoute$1();
		const documentInfo = shallowRef(null);
		const slide = shallowRef(null);
		const slideNumber = ref(0);
		const animationStep = ref(0);
		const blank = ref(null);
		const triggerPlayback = shallowRef([]);
		const transitionPlayback = shallowRef(null);
		const ink = shallowRef([]);
		const draftStroke = shallowRef(null);
		const laserPoint = shallowRef(null);
		const projectorStageShell = ref(null);
		const projectorSlideHost = ref(null);
		const projectorTransitionUnderlay = ref(null);
		const projectorFlashOverlay = ref(null);
		const channelError = ref("");
		const fontWarnings = ref([]);
		const fontsReady = ref(false);
		const slideCountLabel = computed(() => `${slideNumber.value} / ${documentInfo.value?.totalSlides || 0}`);
		function stopAllAudio() {
			for (const audio of (void 0).querySelectorAll("audio")) audio.pause();
		}
		useSlideTransition({
			slide,
			shell: projectorStageShell,
			incomingHost: projectorSlideHost,
			underlay: projectorTransitionUnderlay,
			flash: projectorFlashOverlay,
			playback: transitionPlayback
		});
		watch([
			slide,
			slideNumber,
			animationStep,
			blank,
			triggerPlayback,
			fontsReady
		], () => {
			if (slide.value && fontsReady.value);
		}, { flush: "post" });
		return (_ctx, _push, _parent, _attrs) => {
			const _component_SlideStage = SlideStage_default;
			_push(`<main${ssrRenderAttrs(mergeProps({
				class: ["projector-window", { "is-projector-blank": unref(blank) }],
				"data-projector-window": ""
			}, _attrs))}>`);
			if (unref(slide) && unref(documentInfo) && unref(fontsReady) && !unref(blank)) {
				_push(`<!--[--><div class="slide-transition-shell projector-transition-shell"><div class="${ssrRenderClass([{
					"is-pulling": unref(slide).transition?.effect === "pull",
					"is-splitting-in": unref(slide).transition?.effect === "split" && unref(slide).transition.direction === "in"
				}, "slide-transition-underlay"])}" aria-hidden="true"></div><div class="slide-host"${ssrRenderAttr("data-slide-number", unref(slideNumber))} style="${ssrRenderStyle({ "--slide-ratio": String(unref(documentInfo).width / unref(documentInfo).height) })}">`);
				_push(ssrRenderComponent(_component_SlideStage, {
					class: "projector-stage",
					style: { "--slide-ratio": String(unref(documentInfo).width / unref(documentInfo).height) },
					slide: unref(slide),
					width: unref(documentInfo).width,
					height: unref(documentInfo).height,
					number: unref(slideNumber),
					"animation-step": unref(animationStep),
					"show-animations": unref(documentInfo).showAnimations,
					"show-narration": unref(documentInfo).showNarration,
					"trigger-playback": unref(triggerPlayback),
					onStopAllAudio: stopAllAudio,
					thumbnail: "",
					"playback-preview": "",
					"data-slide-id": unref(slide).id,
					"data-slide-number": unref(slideNumber),
					"data-animation-step": unref(animationStep)
				}, null, _parent));
				_push(`<svg class="ink-layer projector-ink-layer"${ssrRenderAttr("viewBox", `0 0 ${unref(documentInfo).width} ${unref(documentInfo).height}`)} preserveAspectRatio="none" aria-hidden="true"><!--[-->`);
				ssrRenderList(unref(ink), (stroke, index) => {
					_push(`<polyline${ssrRenderAttr("points", stroke.points)} fill="none"${ssrRenderAttr("stroke", stroke.tool === "pen" ? "#ef4444" : "#fde047")}${ssrRenderAttr("stroke-width", stroke.tool === "pen" ? 3 : 16)}${ssrRenderAttr("opacity", stroke.tool === "pen" ? 1 : .4)} stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke"></polyline>`);
				});
				_push(`<!--]-->`);
				if (unref(draftStroke)) _push(`<polyline${ssrRenderAttr("points", unref(draftStroke).points)} fill="none"${ssrRenderAttr("stroke", unref(draftStroke).tool === "pen" ? "#ef4444" : "#fde047")}${ssrRenderAttr("stroke-width", unref(draftStroke).tool === "pen" ? 3 : 16)}${ssrRenderAttr("opacity", unref(draftStroke).tool === "pen" ? 1 : .4)} stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke"></polyline>`);
				else _push(`<!---->`);
				if (unref(laserPoint)) _push(`<circle${ssrRenderAttr("cx", unref(laserPoint).x)}${ssrRenderAttr("cy", unref(laserPoint).y)}${ssrRenderAttr("r", unref(documentInfo).width / 120)} fill="#ef4444" opacity=".9"></circle>`);
				else _push(`<!---->`);
				_push(`</svg></div><div class="flash-transition-overlay" aria-hidden="true"></div></div><div class="projector-toolbar" aria-label="Projector controls"><button class="projector-fullscreen-button">Fullscreen</button><span>${ssrInterpolate(unref(slideCountLabel))} · ← → to navigate · B/W blank · Esc close</span><!--[-->`);
				ssrRenderList(unref(fontWarnings), (warning) => {
					_push(`<span class="warning-pill" role="status">${ssrInterpolate(warning)}</span>`);
				});
				_push(`<!--]--></div><!--]-->`);
			} else if (unref(blank)) _push(`<div class="${ssrRenderClass([`is-${unref(blank)}`, "projector-blank"])}"${ssrRenderAttr("aria-label", `${unref(blank)} screen`)}></div>`);
			else if (unref(channelError)) _push(`<p class="projector-message" role="alert">${ssrInterpolate(unref(channelError))}</p>`);
			else _push(`<p class="projector-message" data-projector-waiting>${ssrInterpolate(unref(documentInfo)?.name || "Waiting for the presenter…")}</p>`);
			_push(`</main>`);
		};
	}
});
//#endregion
//#region app/pages/projector.vue
var _sfc_setup = projector_vue_vue_type_script_setup_true_lang_default.setup;
projector_vue_vue_type_script_setup_true_lang_default.setup = (props, ctx) => {
	const ssrContext = useSSRContext();
	(ssrContext.modules || (ssrContext.modules = /* @__PURE__ */ new Set())).add("pages/projector.vue");
	return _sfc_setup ? _sfc_setup(props, ctx) : void 0;
};
var projector_default = projector_vue_vue_type_script_setup_true_lang_default;

export { projector_default as default };
//# sourceMappingURL=projector-CqXGBXjQ.mjs.map
