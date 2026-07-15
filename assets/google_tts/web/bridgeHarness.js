(function () {
	"use strict";

	let currentSessionId = null;
	let currentOutputGain = 1;
	let lastChunkAt = 0;
	let stopped = false;
	let initPromise = null;
	let cachedEngine = null;
	const engineLoadErrors = [];
	const engineDiscoveryPollMs = 50;
	const engineDiscoveryTimeoutMs = 30000;
	const firstAudioPacketSamples = 48;
	const steadyAudioPacketSamples = 240;
	const synthesisIdlePollMs = 5;
	const synthesisIdleTailMs = 25;
	let emittedAudioPackets = 0;
	let pendingAudioBuffers = [];
	let pendingAudioSampleCount = 0;
	const messageListeners = [];


	function rememberEngineLoadError(value) {
		const message = value && value.message ? value.message : String(value || "Unknown browser engine error");
		engineLoadErrors.push(message);
		if (engineLoadErrors.length > 8) {
			engineLoadErrors.shift();
		}
	}

	window.addEventListener("error", (event) => {
		if (event && event.target && event.target.src) {
			rememberEngineLoadError(`Failed to load ${event.target.src}`);
		} else if (event && event.message) {
			rememberEngineLoadError(event.message);
		}
	}, true);

	window.addEventListener("unhandledrejection", (event) => {
		rememberEngineLoadError(event && event.reason ? event.reason : "Unhandled promise rejection");
	});

	function emit(message) {
		if (!message || !currentSessionId) {
			return;
		}
		message.sessionId = currentSessionId;
		window.googleTtsForSonarpadBridge(JSON.stringify(message));
	}

	function dispatchChromeMessage(message, callback) {
		const run = async () => {
			let response = { result: "stubbed" };
			for (const listener of messageListeners) {
				let listenerResponse;
				const maybePromise = listener(message, { id: "google-tts-for-nvda" }, (value) => {
					listenerResponse = value;
				});
				if (maybePromise && typeof maybePromise.then === "function") {
					listenerResponse = await maybePromise;
				}
				if (listenerResponse !== undefined) {
					response = listenerResponse;
				}
			}
			if (callback) {
				callback(response);
			}
			return response;
		};
		return run();
	}

	const chromeApi = {};
	chromeApi.runtime = {
		onMessage: {
			addListener(listener) {
				messageListeners.push(listener);
			},
		},
		sendMessage(...args) {
			const message = typeof args[0] === "string" ? args[1] : args[0];
			const callback = args.find((arg) => typeof arg === "function");
			return dispatchChromeMessage(message, callback);
		},
		getURL(path) {
			return `/${path.replace(/^\/+/, "")}`;
		},
		getPlatformInfo() {
			return Promise.resolve({ os: "win", arch: "x86-64", nacl_arch: "x86-64" });
		},
		onInstalled: { addListener() {} },
		onStartup: { addListener() {} },
	};
	chromeApi.storage = {
		local: {
			_store: {},
			async get(key) {
				if (typeof key === "string") {
					return { [key]: this._store[key] };
				}
				return { ...this._store };
			},
			async set(values) {
				Object.assign(this._store, values);
			},
		},
	};
	chromeApi.ttsEngine = {
		LanguageInstallStatus: {
			INSTALLED: "installed",
			NOT_INSTALLED: "notInstalled",
			INSTALLING: "installing",
		},
		TtsClientSource: { CHROMEFEATURE: "chrome_feature" },
		updateLanguage() {},
		updateVoices() {},
		onSpeak: { addListener() {} },
		onStop: { addListener() {} },
		onPause: { addListener() {} },
		onResume: { addListener() {} },
		onInstallLanguageRequest: { addListener() {} },
		onLanguageStatusRequest: { addListener() {} },
		onUninstallLanguageRequest: { addListener() {} },
	};
	chromeApi.offscreen = {
		Reason: { AUDIO_PLAYBACK: "AUDIO_PLAYBACK", USER_MEDIA: "USER_MEDIA" },
		async hasDocument() { return true; },
		async createDocument() {},
		async closeDocument() {},
	};
	window.chrome = chromeApi;

	class FakeAudioContext {
		constructor(options) {
			this.sampleRate = options && options.sampleRate ? options.sampleRate : 24000;
			this.destination = {};
			this.audioWorklet = {
				addModule: async () => undefined,
			};
		}

		createGain() {
			return {
				gain: { value: 1 },
				connect() {},
			};
		}

		async resume() {}
		async suspend() {}
	}

	function outputGainFromPayload(payload) {
		const gain = Number(payload && payload.outputGain);
		if (!Number.isFinite(gain)) {
			return 1;
		}
		return Math.max(0, Math.min(2, gain));
	}

	function buffersToPcmBase64(buffers, sampleCount) {
		const bytes = new Uint8Array(sampleCount * 2);
		const view = new DataView(bytes.buffer);
		let outputIndex = 0;
		for (const buffer of buffers) {
			for (let inputIndex = 0; inputIndex < buffer.length; inputIndex++) {
				const sample = Math.max(-1, Math.min(1, buffer[inputIndex] * currentOutputGain));
				view.setInt16(outputIndex * 2, sample < 0 ? sample * 0x8000 : sample * 0x7fff, true);
				outputIndex++;
			}
		}
		let binary = "";
		const chunkSize = 0x8000;
		for (let index = 0; index < bytes.length; index += chunkSize) {
			binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
		}
		return btoa(binary);
	}

	function resetAudioQueue() {
		pendingAudioBuffers = [];
		pendingAudioSampleCount = 0;
		emittedAudioPackets = 0;
	}

	function flushAudioQueue() {
		if (!pendingAudioSampleCount || stopped) {
			resetAudioQueue();
			return;
		}
		emit({
			type: "audio",
			sampleRate: 24000,
			data: buffersToPcmBase64(pendingAudioBuffers, pendingAudioSampleCount),
		});
		pendingAudioBuffers = [];
		pendingAudioSampleCount = 0;
		emittedAudioPackets++;
	}

	function queueAudio(samples) {
		pendingAudioBuffers.push(samples.slice());
		pendingAudioSampleCount += samples.length;
		const packetSamples = emittedAudioPackets === 0 ? firstAudioPacketSamples : steadyAudioPacketSamples;
		if (pendingAudioSampleCount >= packetSamples) {
			flushAudioQueue();
		}
	}

	class FakeAudioWorkletNode {
		constructor() {
			this.port = {
				onmessage: null,
				postMessage(message) {
					if (!message || message.command !== "addBuffer" || !message.buffer || stopped) {
						return;
					}
					const samples = message.buffer instanceof Float32Array
						? message.buffer
						: new Float32Array(message.buffer);
					lastChunkAt = performance.now();
					queueAudio(samples);
				},
			};
		}

		connect() {}
		disconnect() {}
	}

	window.AudioContext = FakeAudioContext;
	window.webkitAudioContext = FakeAudioContext;
	window.AudioWorkletNode = FakeAudioWorkletNode;

	async function waitForSynthesisIdle(timeoutMs, idleMs) {
		const startedAt = performance.now();
		let sawAudio = false;
		while (performance.now() - startedAt < timeoutMs) {
			await new Promise((resolve) => setTimeout(resolve, synthesisIdlePollMs));
			if (stopped) {
				return;
			}
			if (lastChunkAt > 0) {
				sawAudio = true;
			}
			if (sawAudio && performance.now() - lastChunkAt >= idleMs) {
				return;
			}
		}
		throw new Error("Timed out waiting for Chrome TTS audio.");
	}

	function isTtsEngine(value) {
		return value
			&& typeof value === "object"
			&& typeof value.onSpeak === "function"
			&& typeof value.init === "function"
			&& typeof value.onStop === "function";
	}

	function getTtsEngine() {
		if (isTtsEngine(cachedEngine)) {
			return cachedEngine;
		}
		if (isTtsEngine(window.Vh)) {
			cachedEngine = window.Vh;
			return cachedEngine;
		}
		if (isTtsEngine(window.Uh)) {
			cachedEngine = window.Uh;
			return cachedEngine;
		}
		for (const key of Object.getOwnPropertyNames(window)) {
			try {
				const value = window[key];
				if (isTtsEngine(value)) {
					cachedEngine = value;
					return cachedEngine;
				}
			} catch (_) {}
		}
		return null;
	}

	function engineStatus() {
		return {
			ready: Boolean(getTtsEngine()),
			documentReadyState: document.readyState,
			hasBindingsFactory: typeof window.loadWasmTtsBindings === "function",
			errors: engineLoadErrors.slice(-5),
		};
	}

	async function waitForTtsEngine(timeoutMs) {
		const effectiveTimeout = Number.isFinite(Number(timeoutMs))
			? Math.max(1000, Number(timeoutMs))
			: engineDiscoveryTimeoutMs;
		const startedAt = performance.now();
		while (performance.now() - startedAt < effectiveTimeout) {
			const engine = getTtsEngine();
			if (engine) {
				return engine;
			}
			await new Promise((resolve) => setTimeout(resolve, engineDiscoveryPollMs));
		}
		const status = engineStatus();
		throw new Error(`Chrome WASM TTS engine was not loaded. status=${JSON.stringify(status)}`);
	}

	window.googleTtsForSonarpadEngineStatus = engineStatus;
	window.googleTtsForSonarpadWaitForEngine = async function googleTtsForSonarpadWaitForEngine(timeoutMs) {
		await waitForTtsEngine(timeoutMs);
		return engineStatus();
	};

	async function ensureEngineInitialized() {
		const engine = await waitForTtsEngine(engineDiscoveryTimeoutMs);
		if (!initPromise) {
			initPromise = Promise.resolve(engine.init("google-tts-for-nvda")).catch((error) => {
				initPromise = null;
				throw error;
			});
		}
		await initPromise;
	}

	async function stopActiveSynthesis() {
		stopped = true;
		resetAudioQueue();
		const engine = getTtsEngine();
		if (engine && typeof engine.onStop === "function") {
			await engine.onStop();
		}
	}

	window.googleTtsForSonarpadStop = async function googleTtsForSonarpadStop() {
		const sessionId = currentSessionId;
		await stopActiveSynthesis();
		if (currentSessionId === sessionId) {
			currentSessionId = null;
		}
	};

	window.googleTtsForSonarpadPreload = async function googleTtsForSonarpadPreload(payload) {
		currentSessionId = payload.sessionId;
		currentOutputGain = 0;
		lastChunkAt = 0;
		stopped = false;
		resetAudioQueue();
		await ensureEngineInitialized();
		const engine = getTtsEngine();
		if (!engine) {
			throw new Error("Chrome WASM TTS engine was not loaded.");
		}
		await engine.onSpeak("", {
			voiceName: payload.voiceName,
			lang: payload.lang,
			rate: 1,
			pitch: 1,
			volume: 0,
		});
		if (currentSessionId === payload.sessionId) {
			currentSessionId = null;
		}
		return { success: true, preloaded: true };
	};

	window.googleTtsForSonarpadSpeak = async function googleTtsForSonarpadSpeak(payload) {
		try {
			await ensureEngineInitialized();
			const engine = getTtsEngine();
			if (!engine) {
				throw new Error("Chrome WASM TTS engine was not loaded.");
			}
			if (currentSessionId) {
				await stopActiveSynthesis();
			}
			const sessionId = payload.sessionId;
			currentSessionId = sessionId;
			currentOutputGain = outputGainFromPayload(payload);
			lastChunkAt = 0;
			stopped = false;
			resetAudioQueue();
			emit({ type: "started", rate: payload.rate, pitch: payload.pitch, volume: payload.volume });
			await engine.onSpeak(payload.text, {
				voiceName: payload.voiceName,
				lang: payload.lang,
				rate: payload.rate,
				pitch: payload.pitch,
				volume: payload.volume,
			});
			await waitForSynthesisIdle(120000, synthesisIdleTailMs);
			flushAudioQueue();
			emit({ type: "done" });
			if (currentSessionId === sessionId) {
				currentSessionId = null;
			}
			return { success: true };
		} catch (error) {
			emit({ type: "error", message: error && error.message ? error.message : String(error) });
			if (currentSessionId === payload.sessionId) {
				currentSessionId = null;
			}
			throw error;
		}
	};
})();

