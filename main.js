/**
 * @file main.js
 * @description 音感練習：上傳標準音與錄音檔，比對音高並顯示燈號。
 */

const audioContext = new (window.AudioContext || window.webkitAudioContext)();

/** @type {HTMLInputElement | null} */
const refInput = document.getElementById("ref-audio");
/** @type {HTMLInputElement | null} */
const userInput = document.getElementById("user-audio");
/** @type {HTMLInputElement | null} */
const pianoMasterInput = document.getElementById("piano-master");
/** @type {HTMLInputElement | null} */
const refDelayInput = document.getElementById("ref-delay");
/** @type {HTMLInputElement | null} */
const userDelayInput = document.getElementById("user-delay");
/** @type {HTMLInputElement | null} */
const refVolumeInput = document.getElementById("ref-volume");
/** @type {HTMLInputElement | null} */
const userVolumeInput = document.getElementById("user-volume");
/** @type {HTMLSpanElement | null} */
const refVolumeValEl = document.getElementById("ref-volume-val");
/** @type {HTMLSpanElement | null} */
const userVolumeValEl = document.getElementById("user-volume-val");
/** @type {HTMLButtonElement | null} */
const analyzeBtn = document.getElementById("analyze-btn");
/** @type {HTMLDivElement | null} */
const accuracyTextEl = document.getElementById("accuracy-text");
/** @type {HTMLCanvasElement | null} */
const pitchCanvas = document.getElementById("pitch-canvas");
/** @type {HTMLInputElement | null} */
const timelineEl = document.getElementById("timeline");
/** @type {HTMLSpanElement | null} */
const timeCurrentEl = document.getElementById("time-current");
/** @type {HTMLSpanElement | null} */
const timeTotalEl = document.getElementById("time-total");
/** @type {HTMLDivElement | null} */
const phraseBarEl = document.getElementById("phrase-bar");
/** @type {HTMLButtonElement | null} */
const abSetABtn = document.getElementById("ab-set-a");
/** @type {HTMLButtonElement | null} */
const abSetBBtn = document.getElementById("ab-set-b");
/** @type {HTMLButtonElement | null} */
const abClearBtn = document.getElementById("ab-clear");
/** @type {HTMLSpanElement | null} */
const abALabelEl = document.getElementById("ab-a-label");
/** @type {HTMLSpanElement | null} */
const abBLabelEl = document.getElementById("ab-b-label");
/** @type {HTMLInputElement | null} */
const abEnableCheckbox = document.getElementById("ab-enable");
/** @type {HTMLInputElement | null} */
const abGapInput = document.getElementById("ab-gap");
/** @type {HTMLButtonElement | null} */
const startRecordBtn = document.getElementById("start-record-btn");
/** @type {HTMLButtonElement | null} */
const stopRecordBtn = document.getElementById("stop-record-btn");
/** @type {HTMLDivElement | null} */
const pendingActionsEl = document.getElementById("pending-actions");
/** @type {HTMLButtonElement | null} */
const previewPlayBtn = document.getElementById("preview-play");
/** @type {HTMLButtonElement | null} */
const useAndCompareBtn = document.getElementById("use-and-compare");
/** @type {HTMLButtonElement | null} */
const saveAsBtn = document.getElementById("save-as");
/** @type {HTMLButtonElement | null} */
const deletePendingBtn = document.getElementById("delete-pending");
/** @type {HTMLButtonElement | null} */
const startNewTakeBtn = document.getElementById("start-new-take");
/** @type {HTMLDivElement | null} */
const pendingHintEl = document.getElementById("pending-hint");
/** @type {HTMLButtonElement | null} */
const clearSettingsBtn = document.getElementById("clear-current-settings");
/** @type {HTMLSelectElement | null} */
const takeSelectorEl = document.getElementById("take-selector");
/** @type {HTMLSelectElement | null} */
const recordModeEl = document.getElementById("record-mode");
/** @type {HTMLDivElement | null} */
const deviceWarningEl = document.getElementById("device-warning");
/** @type {HTMLButtonElement | null} */
const recordBtn = document.getElementById("record-btn");
/** @type {HTMLDivElement | null} */
const recordStatusEl = document.getElementById("record-status");
/** @type {HTMLDivElement | null} */
const segmentScoresEl = document.getElementById("segment-scores");
/** @type {HTMLDivElement | null} */
const pitchTooltipEl = document.getElementById("pitch-tooltip");
/** @type {HTMLButtonElement | null} */
const rmsAlignBtn = document.getElementById("rmsAlignBtn");
/** @type {HTMLSpanElement | null} */
const rmsAlignStatusEl = document.getElementById("rmsAlignStatus");
/** @type {HTMLButtonElement | null} */
const replayBtn = document.getElementById("replay-btn");

// 更換標準音 / 使用者音檔時，重置分析與播放狀態，避免繼續使用舊 buffer
if (refInput) {
  refInput.addEventListener("change", () => {
    invalidateAnalysisState();
  });
}
if (userInput) {
  userInput.addEventListener("change", () => {
    invalidateAnalysisState();
  });
}

// 更換鋼琴伴奏 Master 時，同樣需要重置分析/播放狀態
if (pianoMasterInput) {
  pianoMasterInput.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) {
      masterPianoBuffer = null;
      invalidateAnalysisState();
      return;
    }
    try {
      const buf = await readFileAsArrayBuffer(file);
      const decoded = await decodeAudio(buf);
      masterPianoBuffer = decoded;
      invalidateAnalysisState();
    } catch (err) {
      console.error("鋼琴伴奏解碼失敗", err);
      alert("鋼琴伴奏解碼失敗，請確認檔案格式是否支援。");
    }
  });
}

/**
 * 供 tooltip 使用：上次繪圖的偏差資料與版面（Deviation View）
 * @type {{ deviations: (number|null)[]; len: number; paddingLeft: number; w: number; durationSec: number } | null}
 */
let lastPitchDrawData = null;
/**
 * 背景音樂 1～4 的檔案輸入
 * @type {(HTMLInputElement | null)[]}
 */
const bgMusicInputs = [
  document.getElementById("bg-music-1"),
  document.getElementById("bg-music-2"),
  document.getElementById("bg-music-3"),
  document.getElementById("bg-music-4"),
];
const bgNameEls = [
  document.getElementById("bg-name-1"),
  document.getElementById("bg-name-2"),
  document.getElementById("bg-name-3"),
  document.getElementById("bg-name-4"),
];
const bgDelayInputs = [
  document.getElementById("bg-delay-1"),
  document.getElementById("bg-delay-2"),
  document.getElementById("bg-delay-3"),
  document.getElementById("bg-delay-4"),
];
const bgVolumeInputs = [
  document.getElementById("bg-volume-1"),
  document.getElementById("bg-volume-2"),
  document.getElementById("bg-volume-3"),
  document.getElementById("bg-volume-4"),
];
const bgVolumeValEls = [
  document.getElementById("bg-volume-val-1"),
  document.getElementById("bg-volume-val-2"),
  document.getElementById("bg-volume-val-3"),
  document.getElementById("bg-volume-val-4"),
];
/** @type {HTMLDivElement | null} */
const bgMusicStatusEl = document.getElementById("bg-music-status");
/** @type {HTMLInputElement | null} */
const bgAutoPlayCheckbox = document.getElementById("bg-auto-play");
/** @type {HTMLButtonElement | null} */
const bgPlayBtn = document.getElementById("bg-play-btn");

/**
 * 背景音樂最多軌數
 * @type {number}
 */
const MAX_BG_TRACKS = 4;

/**
 * 已解碼的背景音樂 buffer，依序對應背景音樂 1～4，未載入為 null
 * @type {(AudioBuffer | null)[]}
 */
let bgBuffers = [null, null, null, null];

/**
 * 鋼琴伴奏 Master buffer（作為全系統唯一時間基準的最高優先 Master Track）
 * @type {AudioBuffer | null}
 */
let masterPianoBuffer = null;

/** @type {AudioBufferSourceNode | null} */
let accompanimentNode = null;
/** @type {GainNode | null} */
let accompanimentGain = null;

/**
 * 各軌上傳的檔名，用於顯示
 * @type {string[]}
 */
let bgFileNames = ["", "", "", ""];

/**
 * 目前正在播放的背景音樂 source 節點（用於停止）
 * @type {AudioBufferSourceNode[]}
 */
let bgSourceNodes = [];

/**
 * 同步播放時 ref / user 的 GainNode（停止時需 disconnect）
 * @type {GainNode | null}
 */
let refGainNode = null;
/** @type {GainNode | null} */
let userGainNode = null;

/**
 * 背景音樂各軌的 GainNode（停止時需 disconnect）
 * @type {GainNode[]}
 */
let bgTrackGainNodes = [];

/**
 * 背景音樂各軌相對 Master Track 的自動 RMS offset（秒）
 * @type {number[]}
 */
let bgAutoOffsetsSec = [0, 0, 0, 0];

/**
 * 背景音樂總音量（0~1），避免過大
 * @type {GainNode}
 */
let bgGainNode = null;

/**
 * 背景音樂是否正在播放
 * @type {boolean}
 */
let isBgPlaying = false;

/**
 * 使用麥克風錄音時暫存的 AudioBuffer
 * @type {AudioBuffer | null}
 */
let recordedUserAudio = null;

/**
 * 分析完成後暫存，供同步播放與播放線動畫使用
 * @type {{ refBuffer: AudioBuffer; userBuffer: AudioBuffer; refPitches: number[]; userPitches: number[]; durationSec: number; refOffsetSec?: number; userOffsetSec?: number } | null}
 */
let lastPlaybackData = null;

/** @type {"v1" | "v2"} 比對模式，v2 時使用 v2PlaybackData */
let compareMode = "v1";
/** @type {{ refBuffer: AudioBuffer; accBuffer?: AudioBuffer; userBuffer: AudioBuffer; durationSec: number } | null} v2 比對完成後的播放資料 */
let v2PlaybackData = null;

/** @type {AudioBufferSourceNode | null} */
let refSourceNode = null;
/** @type {AudioBufferSourceNode | null} */
let userSourceNode = null;
/** 目前播放中的 ref/user/acc source，stopPlayback 與 reset 時會 stop 並清空 */
/** @type {AudioBufferSourceNode | null} */
let activeRefSource = null;
/** @type {AudioBufferSourceNode | null} */
let activeUserSource = null;
/** @type {AudioBufferSourceNode | null} */
let activeAccSource = null;
/** @type {number} */
let playbackAnimationId = 0;
/** @type {boolean} */
let isPlaying = false;
/**
 * 目前播放起點偏移（秒），用於時間軸拖曳 seek
 * @type {number}
 */
let playbackOffsetSec = 0;
/** @type {boolean} */
let updatingTimelineFromPlayback = false;
/**
 * Master Timeline 的起點（AudioContext.currentTime 座標）
 * @type {number}
 */
let masterStartTime = 0;
/**
 * 舊名稱相容（內容時間軸起點），等同於 masterStartTime
 * @type {number}
 */
let contentStartTime = 0;

/**
 * A/B Loop 狀態
 * @type {{ enabled: boolean; aSec: number | null; bSec: number | null; gapSec: number }}
 */
const abLoop = {
  enabled: false,
  aSec: null,
  bSec: null,
  gapSec: 0.3,
};

/**
 * A/B Loop 跳轉鎖定，避免同一幀重複觸發
 * @type {boolean}
 */
let abLoopJumping = false;

/** @type {MediaStream | null} */
let mediaStream = null;
/** @type {ScriptProcessorNode | null} */
let scriptNode = null;
/** @type {Float32Array[]} */
let recordedChunks = [];
/** @type {boolean} */
let isRecording = false;

// 專業錄音（MediaRecorder）狀態
/** @type {MediaRecorder | null} */
let proMediaRecorder = null;
/** @type {Blob[]} */
let proRecordingChunks = [];
/** @type {boolean} */
let isProRecording = false;
/** @type {AudioBufferSourceNode | null} */
let previewSourceNode = null;
/** @type {number} */
let recordingStartTime = 0;
/** @type {number | null} */
let recordingTimerInterval = null;

/**
 * 錄音 UI 狀態："idle" | "recording" | "pending"
 * @typedef {"idle" | "recording" | "pending"} RecUiState
 */
/** @type {RecUiState} */
let recUiState = "idle";

/**
 * 單次錄音完成後的暫存 pending 錄音（尚未決定 Use / Save / Discard）
 * @type {{ buffer: AudioBuffer; blob?: Blob; engine: "mr" | "wav"; createdAt: number; pitches?: number[]; durationSec?: number } | null}
 */
let pendingRecording = null;

/**
 * 允許直接「使用並比對」的最長秒數（超過則只能另存）
 */
const MAX_COMPARE_SEC = 30;

/**
 * 錄音引擎類型："mr" | "wav" | "none"
 * - "mr": MediaRecorder
 * - "wav": WebAudio PCM → WAV fallback
 */
let recEngine = "none";

/**
 * 單次錄音最大秒數（依錄音引擎調整，產品第一階段兩者皆為 360）
 * - MediaRecorder: 360 秒（約 6 分鐘）
 * - WAV fallback: 360 秒（記憶體占用較高，但先以產品需求為主）
 */
let MAX_RECORDING_SEC = 360;

/**
 * 專業錄音多 take 管理
 * @type {{ id: number; buffer: AudioBuffer; pitches: number[]; createdAt: number }[]}
 */
let recordings = [];
let nextRecordingId = 1;

// WAV fallback state（ScriptProcessor PCM → WAV Blob → AudioBuffer）
/** @type {MediaStream | null} */
let wavStream = null;
/** @type {MediaStreamAudioSourceNode | null} */
let wavSource = null;
/** @type {ScriptProcessorNode | null} */
let wavProcessor = null;
/** @type {Float32Array[]} */
let wavChunks = [];
/** @type {number} */
let wavSampleRate = 48000;
/** @type {number} */
let wavNumChannels = 1;

/**
 * 使用 MediaRecorder 進行「專業錄音」：
 * - 共用同一個 AudioContext
 * - 不改 master timeline / 播放架構
 */
async function initMic() {
  if (mediaStream) return mediaStream;
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    alert("此瀏覽器不支援麥克風錄音。");
    throw new Error("no getUserMedia");
  }
  // 優先關閉瀏覽器端 AGC / Noise Suppression / Echo Cancellation，改由後端穩定化
  const constraints = {
    audio: {
      echoCancellation: false,
      noiseSuppression: false,
      autoGainControl: false,
      channelCount: 1,
    },
  };
  try {
    mediaStream = await navigator.mediaDevices.getUserMedia(constraints);
  } catch (_) {
    // 某些環境可能不支援上述欄位，退回最基本的 { audio: true }
    mediaStream = await navigator.mediaDevices.getUserMedia({ audio: true });
  }
  return mediaStream;
}

/**
 * 停止僅伴奏播放（不影響主播放引擎）
 */
function stopAccompanimentOnly() {
  if (accompanimentNode) {
    try {
      accompanimentNode.stop();
    } catch (_) {
      // ignore
    }
    accompanimentNode = null;
  }
}

/**
 * 僅播放鋼琴伴奏 Master（不經過錄音節點，避免被錄進去）
 * @param {number} startSec
 */
function playAccompanimentOnly(startSec = 0) {
  if (!masterPianoBuffer) return;

  stopAccompanimentOnly();

  if (!accompanimentGain) {
    accompanimentGain = audioContext.createGain();
    accompanimentGain.gain.value = 1.0;
    accompanimentGain.connect(audioContext.destination);
  }

  const src = audioContext.createBufferSource();
  src.buffer = masterPianoBuffer;
  src.connect(accompanimentGain);
  src.start(0, Math.max(0, startSec));
  accompanimentNode = src;
}

/**
 * 將 Float32 PCM 轉為 16-bit PCM
 * @param {Float32Array} float32
 * @returns {Int16Array}
 */
function floatTo16BitPCM(float32) {
  const len = float32.length;
  const out = new Int16Array(len);
  for (let i = 0; i < len; i += 1) {
    let s = Math.max(-1, Math.min(1, float32[i]));
    out[i] = s < 0 ? s * 0x8000 : s * 0x7fff;
  }
  return out;
}

/**
 * 將多個 mono Float32Array chunk 編碼成 16-bit PCM WAV（單聲道）
 * @param {Float32Array[]} chunks
 * @param {number} sampleRate
 * @returns {Blob}
 */
function encodeWav16Mono(chunks, sampleRate) {
  if (!chunks.length) return new Blob([], { type: "audio/wav" });

  let totalLength = 0;
  for (const ch of chunks) totalLength += ch.length;

  const pcmFloat = new Float32Array(totalLength);
  let offset = 0;
  for (const ch of chunks) {
    pcmFloat.set(ch, offset);
    offset += ch.length;
  }

  const pcm16 = floatTo16BitPCM(pcmFloat);
  const bytesPerSample = 2;
  const blockAlign = wavNumChannels * bytesPerSample;
  const byteRate = sampleRate * blockAlign;
  const dataSize = pcm16.length * bytesPerSample;

  const buffer = new ArrayBuffer(44 + dataSize);
  const view = new DataView(buffer);
  let p = 0;

  function writeString(s) {
    for (let i = 0; i < s.length; i += 1) {
      view.setUint8(p += 1 - 1, s.charCodeAt(i));
      p += 1;
    }
  }

  writeString("RIFF");
  view.setUint32(p, 36 + dataSize, true);
  p += 4;
  writeString("WAVE");
  writeString("fmt ");
  view.setUint32(p, 16, true);
  p += 4;
  view.setUint16(p, 1, true);
  p += 2;
  view.setUint16(p, wavNumChannels, true);
  p += 2;
  view.setUint32(p, sampleRate, true);
  p += 4;
  view.setUint32(p, byteRate, true);
  p += 4;
  view.setUint16(p, blockAlign, true);
  p += 2;
  view.setUint16(p, bytesPerSample * 8, true);
  p += 2;
  writeString("data");
  view.setUint32(p, dataSize, true);
  p += 4;

  for (let i = 0; i < pcm16.length; i += 1, p += 2) {
    view.setInt16(p, pcm16[i], true);
  }

  return new Blob([buffer], { type: "audio/wav" });
}

/**
 * 啟動 WAV fallback 錄音：ScriptProcessor 收集 PCM，避免回授
 */
async function startWavFallbackRecording() {
  await initMic();
  if (!mediaStream) throw new Error("no mediaStream");

  wavChunks = [];
  wavSampleRate = audioContext.sampleRate;
  wavNumChannels = 1;

  wavStream = mediaStream;
  wavSource = audioContext.createMediaStreamSource(wavStream);

  const bufferSize = 4096;
  wavProcessor = audioContext.createScriptProcessor(
    bufferSize,
    wavNumChannels,
    wavNumChannels
  );
  wavProcessor.onaudioprocess = (e) => {
    const input = e.inputBuffer.getChannelData(0);
    wavChunks.push(new Float32Array(input));
  };

  const zeroGain = audioContext.createGain();
  zeroGain.gain.value = 0;

  wavSource.connect(wavProcessor);
  wavProcessor.connect(zeroGain);
  zeroGain.connect(audioContext.destination);
}

/**
 * 結束 WAV fallback，編碼 WAV 並交給共用 finalizeRecordedBuffer
 */
async function stopWavFallbackRecordingAndFinalize() {
  if (!wavProcessor || !wavSource || !wavStream) return;

  try {
    wavProcessor.disconnect();
  } catch (_) {
    // ignore
  }
  try {
    wavSource.disconnect();
  } catch (_) {
    // ignore
  }

  const wavBlob = encodeWav16Mono(wavChunks, wavSampleRate);
  wavChunks = [];
  wavProcessor = null;
  wavSource = null;
  wavStream = null;

  const arrayBuffer = await wavBlob.arrayBuffer();
  const buffer = await audioContext.decodeAudioData(arrayBuffer);
  await finalizeRecordedBuffer(buffer, "wav");
}

/**
 * 將一段完成的錄音 buffer 納入 takes，更新選單與預覽 UI
 * @param {AudioBuffer} buffer
 * @param {"mr" | "wav"} engine
 */
async function finalizeRecordedBuffer(buffer, engine) {
  // 先做簡單 normalize / limiter，穩定錄音音量
  const stabilized = stabilizeRecordedBuffer(buffer);

  pendingRecording = {
    buffer: stabilized,
    engine,
    createdAt: Date.now(),
    pitches: [],
    durationSec: stabilized ? stabilized.duration : 0,
  };

  // 進入 pending 狀態：顯示預覽與決策按鈕，但不自動播放
  enterPendingState();

  isProRecording = false;
  if (recordingTimerInterval != null) {
    clearInterval(recordingTimerInterval);
    recordingTimerInterval = null;
  }
}

/**
 * 檢查錄音相關能力（secure context / getUserMedia / MediaRecorder）
 */
function getRecordingCapability() {
  return {
    isSecureContext: window.isSecureContext,
    hasGetUserMedia:
      !!(navigator.mediaDevices && navigator.mediaDevices.getUserMedia),
    hasMediaRecorder: typeof window.MediaRecorder !== "undefined",
    userAgent: navigator.userAgent,
  };
}

/**
 * 偵測音訊輸入/輸出裝置，並根據藍牙 / 錄音模式顯示提示
 */
async function checkAudioDevices() {
  if (!navigator.mediaDevices || !navigator.mediaDevices.enumerateDevices) {
    return;
  }

  const devices = await navigator.mediaDevices.enumerateDevices();
  const inputs = devices.filter((d) => d.kind === "audioinput");
  const outputs = devices.filter((d) => d.kind === "audiooutput");

  const hasBluetoothInput = inputs.some(
    (d) => d.label && d.label.toLowerCase().includes("bluetooth")
  );
  const hasBluetoothOutput = outputs.some(
    (d) => d.label && d.label.toLowerCase().includes("bluetooth")
  );

  if (hasBluetoothInput && hasBluetoothOutput) {
    setDeviceWarning(
      "error",
      "⚠ 偵測到藍牙雙向模式，可能影響音準分析。建議改用有線耳機。"
    );
  } else if (recordModeEl && recordModeEl.value === "speaker") {
    setDeviceWarning(
      "warning",
      "⚠ 外放模式可能將伴奏錄入麥克風，影響準確度。"
    );
  } else if (recordModeEl && recordModeEl.value === "headphone") {
    setDeviceWarning(
      "warning",
      "🎧 建議使用耳機以避免伴奏被錄入。"
    );
  }
}

/**
 * 設定錄音裝置/引擎相關提示
 * @param {"warning" | "error" | ""} kind
 * @param {string} text
 */
function setDeviceWarning(kind, text) {
  if (!deviceWarningEl) return;
  deviceWarningEl.className = "device-warning";
  deviceWarningEl.textContent = "";
  if (!kind) return;
  deviceWarningEl.textContent = text;
  deviceWarningEl.classList.add(kind);
}

/**
 * 進入 Idle 狀態：只顯示「開始錄音」，隱藏停止與 pending 區塊
 */
function enterIdleState() {
  recUiState = "idle";

  if (startRecordBtn) {
    startRecordBtn.hidden = false;
    startRecordBtn.disabled = false;
    startRecordBtn.textContent = "開始錄音";
  }
  if (stopRecordBtn) {
    stopRecordBtn.hidden = true;
    stopRecordBtn.disabled = false;
    stopRecordBtn.textContent = "停止錄音";
  }
  if (pendingActionsEl) {
    pendingActionsEl.hidden = true;
  }
   if (pendingHintEl) {
     pendingHintEl.textContent = "";
   }
}

/**
 * 進入 Recording 狀態：只顯示「停止錄音」按鈕
 */
function enterRecordingState() {
  recUiState = "recording";

  if (startRecordBtn) {
    startRecordBtn.hidden = true;
  }
  if (stopRecordBtn) {
    stopRecordBtn.hidden = false;
    stopRecordBtn.disabled = false;
    stopRecordBtn.textContent = "停止錄音";
  }
  if (pendingActionsEl) {
    pendingActionsEl.hidden = true;
  }
  if (pendingHintEl) {
    pendingHintEl.textContent = "";
  }
}

/**
 * 進入 Pending 狀態：顯示「播放預覽 / 使用並比對 / 另存 / 放棄」
 */
function enterPendingState() {
  recUiState = "pending";

  if (startRecordBtn) {
    startRecordBtn.hidden = true;
  }
  if (stopRecordBtn) {
    stopRecordBtn.hidden = true;
  }
  if (pendingActionsEl) {
    pendingActionsEl.hidden = false;
  }

  // 30 秒規則：禁用「使用並比對」，但允許另存/刪除/新錄音
  const dur =
    (pendingRecording && (pendingRecording.durationSec || pendingRecording.buffer?.duration)) ||
    0;
  if (useAndCompareBtn) {
    useAndCompareBtn.disabled = dur > MAX_COMPARE_SEC;
    useAndCompareBtn.title =
      dur > MAX_COMPARE_SEC
        ? `此錄音約 ${Math.round(
            dur
          )} 秒，超過 ${MAX_COMPARE_SEC} 秒，暫不支援直接比對；請用「另存新檔」再用上傳方式比對。`
        : "";
  }
  if (pendingHintEl) {
    pendingHintEl.textContent =
      dur > MAX_COMPARE_SEC
        ? `此錄音約 ${Math.round(
            dur
          )} 秒：可另存新檔，但暫不支援直接比對。`
        : "請選擇：預覽 / 使用並比對 / 另存 / 刪除 / 新錄音";
  }
}

/**
 * 清空 pending 錄音與相關 UI / 預覽
 */
function clearPendingRecording() {
  if (previewSourceNode) {
    try {
      previewSourceNode.stop();
    } catch (_) {
      // ignore
    }
    previewSourceNode = null;
  }

  pendingRecording = null;
  proRecordingChunks = [];
  wavChunks = [];

  enterIdleState();
}

async function startProRecording() {
  if (isProRecording) return;

  const cap = getRecordingCapability();
  if (!cap.isSecureContext || !cap.hasGetUserMedia) {
    return;
  }

  await checkAudioDevices().catch(() => {});

  try {
    await initMic();
  } catch {
    return;
  }

  if (!mediaStream) return;

  // 決定錄音引擎：優先使用 MediaRecorder，失敗或不支援則改用 WAV fallback
  recEngine = "none";
  proRecordingChunks = [];
  proMediaRecorder = null;

  if (cap.hasMediaRecorder) {
    try {
      const mr = createSafeMediaRecorder(mediaStream);
      if (mr) {
        proMediaRecorder = mr;
        recEngine = "mr";
      }
    } catch (e) {
      console.error("建立 MediaRecorder 失敗，將改用 WAV fallback", e);
    }
  }

  if (recEngine === "mr" && proMediaRecorder) {
    MAX_RECORDING_SEC = 360;
    setDeviceWarning(
      "warning",
      "錄音引擎：MediaRecorder（opus/webm 優先，可長錄音）"
    );

    proMediaRecorder.onerror = (e) => {
      console.error("MediaRecorder 錯誤：", e);
      alert(
        "錄音裝置不支援目前格式，請改用有線耳機或重新整理頁面，或改用支援度較高的瀏覽器。"
      );
    };

    proMediaRecorder.ondataavailable = (e) => {
      if (e.data && e.data.size > 0) {
        proRecordingChunks.push(e.data);
      }
    };

    proMediaRecorder.onstop = async () => {
      try {
        const blob = new Blob(proRecordingChunks, { type: "audio/webm" });
        const arrayBuffer = await blob.arrayBuffer();
        const buffer = await audioContext.decodeAudioData(arrayBuffer);
        await finalizeRecordedBuffer(buffer, "mr");
      } catch (e) {
        console.error("錄音解碼失敗", e);
        alert("錄音解碼失敗，請再試一次。");
      }
    };
  } else {
    // MediaRecorder 不可用或建立失敗 → 改用 WAV fallback
    recEngine = "wav";
    MAX_RECORDING_SEC = 360;
    setDeviceWarning(
      "warning",
      "錄音引擎：WAV Fallback（改用 PCM→WAV，相容性較高但較吃記憶體）。"
    );
  }

  isProRecording = true;
  enterRecordingState();

  // 錄音計時器：更新文案並檢查上限
  recordingStartTime = performance.now();
  if (recordingTimerInterval != null) {
    clearInterval(recordingTimerInterval);
    recordingTimerInterval = null;
  }
  recordingTimerInterval = window.setInterval(() => {
    const sec = Math.floor((performance.now() - recordingStartTime) / 1000);

    if (stopRecordBtn) {
      stopRecordBtn.textContent = `停止錄音（${sec}s / ${MAX_RECORDING_SEC}s）`;
    }

    if (sec >= MAX_RECORDING_SEC) {
      if (recordingTimerInterval != null) {
        clearInterval(recordingTimerInterval);
        recordingTimerInterval = null;
      }
      alert("已達錄音上限。");
      // 必須能停得掉（即使 isProRecording 被提前改掉）
      // guard 改由 stopProRecording 內部依 recEngine 決定
      stopProRecording();
    }
  }, 500);

  // 先停掉任何現有播放（避免疊音）
  if (isPlaying) {
    stopPlayback();
  }
  // 若有鋼琴伴奏 Master，錄音一開始就獨立播放伴奏（不依賴已分析狀態）
  if (masterPianoBuffer) {
    playAccompanimentOnly(0);
  }

  if (recEngine === "mr" && proMediaRecorder) {
    proRecordingChunks = [];
    proMediaRecorder.start();
  } else if (recEngine === "wav") {
    await startWavFallbackRecording();
  }
}

async function stopProRecording() {
  // 不用 isProRecording 當硬門檻，避免 timer/狀態不同步導致 stop 失效
  try {
    if (recEngine === "mr") {
      if (proMediaRecorder && proMediaRecorder.state !== "inactive") {
        proMediaRecorder.stop();
      }
    } else if (recEngine === "wav") {
      await stopWavFallbackRecordingAndFinalize();
    }
  } catch (_) {
    // ignore
  }

  // 停止錄音時也停止當前播放（伴奏、ref/user）
  if (isPlaying) {
    stopPlayback();
  }

  // 若有獨立伴奏播放，也一併停止
  stopAccompanimentOnly();

  if (recordingTimerInterval != null) {
    clearInterval(recordingTimerInterval);
    recordingTimerInterval = null;
  }

  recEngine = "none";
}

/**
 * 播放最新錄音的 preview
 * @param {AudioBuffer} buffer
 */
function playPreview(buffer) {
  if (!buffer) return;

  if (previewSourceNode) {
    try {
      previewSourceNode.stop();
    } catch (_) {
      // ignore
    }
  }

  const source = audioContext.createBufferSource();
  source.buffer = buffer;
  source.connect(audioContext.destination);
  source.start();

  previewSourceNode = source;
}

/**
 * 輕量穩定錄音音量：peak normalize 到 0.9，並做簡單 limiter
 * @param {AudioBuffer} buffer
 * @returns {AudioBuffer}
 */
function stabilizeRecordedBuffer(buffer) {
  const numChannels = buffer.numberOfChannels;
  const length = buffer.length;
  const sampleRate = buffer.sampleRate;

  if (numChannels <= 0 || length <= 0) return buffer;

  let peak = 0;
  for (let ch = 0; ch < numChannels; ch += 1) {
    const data = buffer.getChannelData(ch);
    for (let i = 0; i < length; i += 1) {
      const v = Math.abs(data[i]);
      if (v > peak) peak = v;
    }
  }

  if (!Number.isFinite(peak) || peak <= 0) {
    return buffer;
  }

  const target = 0.9;
  const gain = target / peak;

  const out = audioContext.createBuffer(numChannels, length, sampleRate);
  for (let ch = 0; ch < numChannels; ch += 1) {
    const inData = buffer.getChannelData(ch);
    const outData = out.getChannelData(ch);
    for (let i = 0; i < length; i += 1) {
      let v = inData[i] * gain;
      // 簡單 soft limiter：避免硬剪裁
      const abs = Math.abs(v);
      if (abs > 1) {
        const sign = v < 0 ? -1 : 1;
        v = sign * (1 - Math.exp(-abs + 1));
      }
      outData[i] = v;
    }
  }

  return out;
}

/**
 * Playback Snapshot：保存目前 lastPlaybackData / timeline / phrases，供 Undo 使用
 */
let playbackSnapshot = null;

function takePlaybackSnapshot() {
  if (!lastPlaybackData) return;
  playbackSnapshot = {
    lastPlaybackData: {
      ...lastPlaybackData,
      refPitches: Array.isArray(lastPlaybackData.refPitches)
        ? [...lastPlaybackData.refPitches]
        : [],
      userPitches: Array.isArray(lastPlaybackData.userPitches)
        ? [...lastPlaybackData.userPitches]
        : [],
    },
    playbackOffsetSec,
    timelineValue: timelineEl ? timelineEl.value : "0",
    phrases: Array.isArray(phrases) ? phrases.map((p) => ({ ...p })) : [],
  };
}

function restorePlaybackSnapshot() {
  if (!playbackSnapshot) return;

  lastPlaybackData = playbackSnapshot.lastPlaybackData;
  playbackOffsetSec = playbackSnapshot.playbackOffsetSec || 0;
  if (timelineEl) timelineEl.value = playbackSnapshot.timelineValue || "0";
  phrases = Array.isArray(playbackSnapshot.phrases)
    ? playbackSnapshot.phrases.map((p) => ({ ...p }))
    : [];
  renderPhraseBar();

  const dur = lastPlaybackData.durationSec || 0;
  drawPitchCurves(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    undefined,
    dur
  );
  const accuracy = computeAccuracy(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches
  );
  updateUIWithAccuracy(accuracy);
  const segmentAcc = computeSegmentAccuracies(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    SEGMENT_COUNT
  );
  updateSegmentScores(segmentAcc);
  if (timeTotalEl) timeTotalEl.textContent = formatTime(dur);
}

/**
 * 確保 baseline（標準音）已完成最小必要分析，建立 lastPlaybackData
 * - 只要有 refInput 檔案即可自動 analyze 一次
 * @returns {Promise<boolean>}
 */
async function ensureAnalyzedBaseline() {
  // 必須已有標準音檔
  if (!refInput) {
    alert("請先載入標準音檔。");
    return false;
  }
  const fileList = refInput.files;
  if (!fileList || !fileList.length) {
    alert("請先載入標準音檔。");
    return false;
  }

  // 若已有有效的 lastPlaybackData（refPitches 已算好）就直接使用
  if (
    lastPlaybackData &&
    lastPlaybackData.refPitches &&
    lastPlaybackData.refPitches.length
  ) {
    return true;
  }

  try {
    const file = fileList[0];
    const buf = await readFileAsArrayBuffer(file);
    const refAudio = await decodeAudio(buf);

    const refPitches = extractPitchSeries(refAudio);
    const durationSec = refAudio.duration;

    lastPlaybackData = {
      refBuffer: refAudio,
      userBuffer: null,
      refPitches,
      userPitches: [],
      durationSec,
      refOffsetSec: 0,
      userOffsetSec: 0,
    };

    playbackOffsetSec = 0;
    if (timelineEl) timelineEl.value = "0";
    if (timeTotalEl) timeTotalEl.textContent = formatTime(durationSec);

    // 分句偵測：先用 refPitches 即可，user 之後才加入
    phrases = detectPhrases(refPitches, durationSec);
    renderPhraseBar();

    return true;
  } catch (err) {
    console.error("自動分析標準音失敗", err);
    alert("自動分析標準音失敗，請確認檔案格式是否支援。");
    return false;
  }
}

/**
 * 使用 pendingRecording 進行對齊與圖表更新（不自動存成 take）
 */
async function usePendingRecordingAndCompare() {
  const ok = await ensureAnalyzedBaseline();
  if (!ok) return;
  if (!pendingRecording) {
    return;
  }

  // 30 秒 guard：長錄音只能另存，不允許直接比對
  const dur =
    (pendingRecording.durationSec || pendingRecording.buffer?.duration) || 0;
  if (dur > MAX_COMPARE_SEC) {
    alert(
      `此錄音約 ${Math.round(
        dur
      )} 秒，超過 ${MAX_COMPARE_SEC} 秒，暫不支援直接比對。請點「另存新檔」後用上傳方式比對。`
    );
    return;
  }

  // 在覆寫 lastPlaybackData.user* 之前先拍一份快照，供「清除目前設定」還原
  takePlaybackSnapshot();

  // 停止任何當前播放（避免錄音與播放疊在一起）
  if (isPlaying) {
    stopPlayback();
  }

  const refPitches = lastPlaybackData.refPitches;
  if (!refPitches || !refPitches.length) {
    return;
  }

  const userBuffer = pendingRecording.buffer;
  const processedUserBuffer = preprocessUserRecording(userBuffer);

  // 1) pitch 分析（若尚未計算）：沿用既有 pitch pipeline
  let userPitches = pendingRecording.pitches;
  if (!userPitches || !userPitches.length) {
    userPitches = analyzePitchFromBuffer(processedUserBuffer);
    pendingRecording.pitches = userPitches;
  }

  // 2) 與 ref 進行時間對齊（直接使用已計算好的 refPitches）
  const aligned = alignPitchSeries(refPitches, userPitches);

  // 3) 更新 lastPlaybackData 裡的 user 部分（ref 部分保持不動）
  const durationSec = lastPlaybackData.durationSec;
  lastPlaybackData = {
    ...lastPlaybackData,
    userBuffer: processedUserBuffer,
    userPitches: aligned.user,
    durationSec,
    userOffsetSec: aligned.userOffsetSec ?? 0,
  };

  // 4) 重新畫圖與分段分數
  const accuracy = computeAccuracy(lastPlaybackData.refPitches, lastPlaybackData.userPitches);
  const segmentAccuracies = computeSegmentAccuracies(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    SEGMENT_COUNT
  );

  updateUIWithAccuracy(accuracy);
  drawPitchCurves(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    undefined,
    lastPlaybackData.durationSec
  );
  if (segmentScoresEl) {
    updateSegmentScores(segmentAccuracies);
  } else {
    updateSegmentScores(segmentAccuracies);
  }
  if (timeTotalEl) timeTotalEl.textContent = formatTime(lastPlaybackData.durationSec);

  // 5) 重新偵測分句與 phrase bar
  const hasValidRef = lastPlaybackData.refPitches.some(
    (p) => p && Number.isFinite(p) && p > 0
  );
  const phraseSource = hasValidRef ? lastPlaybackData.refPitches : lastPlaybackData.userPitches;
  phrases = detectPhrases(phraseSource, lastPlaybackData.durationSec);
  renderPhraseBar();

  // 自動從 0 秒開始播放 ref + 伴奏 + 使用者錄音
  playbackOffsetSec = 0;
  startSyncedPlayback();
  updateReplayButtonState();

  // 將此錄音存入 recordings，並更新 take selector
  const rec = {
    id: nextRecordingId++,
    buffer: processedUserBuffer,
    pitches: aligned.user,
    createdAt: Date.now(),
  };
  recordings.push(rec);
  if (takeSelectorEl) {
    const option = document.createElement("option");
    option.value = String(rec.id);
    option.textContent = `Take ${rec.id}`;
    takeSelectorEl.appendChild(option);
    takeSelectorEl.value = String(rec.id);
  }
}

/**
 * 錄音前的音訊預處理 hook（第一階段先直接回傳，第二階段再加入處理）
 * @param {AudioBuffer} userBuffer
 * @returns {AudioBuffer}
 */
function preprocessUserRecording(userBuffer) {
  // TODO: 第二階段將在此加入：
  // - 伴奏相位抵消
  // - Noise Gate
  // - 高通濾波
  return userBuffer;
}

/**
 * 從 AudioBuffer 抽取 pitch 序列（沿用既有 pitch pipeline）
 * @param {AudioBuffer} buffer
 * @returns {number[]}
 */
function analyzePitchFromBuffer(buffer) {
  return extractPitchSeries(buffer);
}

/**
 * 專業錄音確認：使用最新一個 take，重新對齊並更新 lastPlaybackData 的 user 部分
 */
async function confirmLatestRecording() {
  if (!lastPlaybackData) {
    alert("請先載入並分析標準音。");
    return;
  }
  if (!recordings.length) {
    alert("目前沒有可用的錄音。");
    return;
  }

  let rec = recordings[recordings.length - 1];
  if (takeSelectorEl && takeSelectorEl.value) {
    const id = Number(takeSelectorEl.value);
    const found = recordings.find((r) => r.id === id);
    if (found) rec = found;
  }

  const refBuffer = lastPlaybackData.refBuffer;
  const processedUserBuffer = preprocessUserRecording(rec.buffer);

  // 1) pitch 分析（若尚未計算）
  if (!rec.pitches.length) {
    rec.pitches = analyzePitchFromBuffer(processedUserBuffer);
  }

  // 2) 與 ref 進行時間對齊（直接使用已計算好的 refPitches）
  const aligned = alignPitchSeries(lastPlaybackData.refPitches, rec.pitches);

  // 3) 更新 lastPlaybackData 裡的 user 部分（ref 部分保持不動）
  const durationSec = lastPlaybackData.durationSec;
  lastPlaybackData = {
    ...lastPlaybackData,
    userBuffer: processedUserBuffer,
    userPitches: aligned.user,
    durationSec,
    userOffsetSec: aligned.userOffsetSec ?? 0,
  };

  // 4) 重新畫圖與分段分數
  const accuracy = computeAccuracy(lastPlaybackData.refPitches, lastPlaybackData.userPitches);
  const segmentAccuracies = computeSegmentAccuracies(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    SEGMENT_COUNT
  );

  updateUIWithAccuracy(accuracy);
  drawPitchCurves(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    undefined,
    lastPlaybackData.durationSec
  );
  updateSegmentScores(segmentAccuracies);
  if (timeTotalEl) timeTotalEl.textContent = formatTime(lastPlaybackData.durationSec);

  // 5) 重新偵測分句與 phrase bar
  const hasValidRef = lastPlaybackData.refPitches.some(
    (p) => p && Number.isFinite(p) && p > 0
  );
  const phraseSource = hasValidRef ? lastPlaybackData.refPitches : lastPlaybackData.userPitches;
  phrases = detectPhrases(phraseSource, lastPlaybackData.durationSec);
  renderPhraseBar();
}

/**
 * 麥克風錄音預設秒數（10 分鐘）
 * @type {number}
 */
const RECORD_SECONDS = 600;

/**
 * 允許的音準誤差（cent）
 * 這裡先抓 ±50 cent，大約半音的 1/2。
 * @type {number}
 */
const CENT_TOLERANCE = 50;

/**
 * 取樣間隔（秒）
 * 例如 0.05 = 每 50ms 抓一次音高，約每秒 20 次。
 * @type {number}
 */
const SAMPLE_INTERVAL_SECONDS = 0.05;

/**
 * 分句偵測參數：句與句之間的最小無聲區（ms）與最小句長（ms）
 */
const PHRASE_GAP_MS = 300;
const PHRASE_MIN_MS = 800;

/**
 * 句段拖曳時允許的最小句長（秒）
 */
const PHRASE_MIN_ADJUST_SEC = 0.3;

/**
 * 分句資料（由 pitch 自動偵測填入）
 * @type {{ id: number; startSec: number; endSec: number }[]}
 */
let phrases = [];

/**
 * 分句拖曳狀態
 * @type {number}
 */
let phraseDragIndex = -1;
/** @type {boolean} */
let phraseDragIsLeft = false;
/** @type {number} */
let phraseDragStartX = 0;
/** @type {number} */
let phraseDragStartStartSec = 0;
/** @type {number} */
let phraseDragStartEndSec = 0;
/** @type {number} */
let phraseDragTimelineWidth = 1;

/**
 * 清除所有現有設定（不刪檔）：停止播放、清比對資料、清圖。不呼叫 restorePlaybackSnapshot。
 */
function resetAllSessionState() {
  console.log("RESET SESSION");

  stopPlayback();
  if (typeof stopPlaybackV2 === "function") stopPlaybackV2();

  activeRefSource = null;
  activeUserSource = null;
  activeAccSource = null;

  lastPlaybackData = null;
  v2PlaybackData = null;
  pendingRecording = null;
  if (pendingActionsEl) pendingActionsEl.hidden = true;

  if (pitchCanvas) {
    const ctx = pitchCanvas.getContext("2d");
    if (ctx) ctx.clearRect(0, 0, pitchCanvas.width, pitchCanvas.height);
  }

  if (typeof resetPlayhead === "function") resetPlayhead();

  updateReplayButtonState();
}

/**
 * 依是否有比對資料更新「重新播放」按鈕的 disabled 狀態
 */
function updateReplayButtonState() {
  if (!replayBtn) return;
  replayBtn.disabled = !(lastPlaybackData || v2PlaybackData);
}

/**
 * 以目前 UI 的延遲設定重新播放：強制重建 source，讀取最新 delay。
 */
function replayWithCurrentSettings() {
  console.log("REPLAY CLICKED");

  stopPlayback();

  if (lastPlaybackData) {
    startSyncedPlayback({
      refBuffer: lastPlaybackData.refBuffer,
      userBuffer: lastPlaybackData.userBuffer,
      accBuffer: masterPianoBuffer || undefined,
      durationSec: lastPlaybackData.durationSec,
    });
  }
}

/**
 * 當更換 ref/user 音檔來源時，需清除舊的分析與播放狀態，避免繼續播放舊 AudioBuffer
 * - 停止目前播放
 * - 清空 lastPlaybackData / RMS offset / 視覺化
 */
function invalidateAnalysisState() {
  stopPlayback();
  lastPlaybackData = null;
  v2PlaybackData = null;
  playbackOffsetSec = 0;
  if (typeof audioOffsetsSec !== "undefined") {
    audioOffsetsSec.ref = 0;
    audioOffsetsSec.user = 0;
  }
   bgAutoOffsetsSec = [0, 0, 0, 0];
  if (rmsAlignStatusEl) rmsAlignStatusEl.textContent = "";
  clearVisualizations();
}

/**
 * 對一段 AudioBuffer 計算簡單 RMS energy envelope
 * @param {AudioBuffer} buffer
 * @param {number} [hopMs] frame hop (毫秒)
 * @param {number} [frameMs] frame 長度 (毫秒)
 * @returns {{ env: number[]; hopSec: number }}
 */
function computeRmsEnvelope(buffer, hopMs = 20, frameMs = 40) {
  const sr = buffer.sampleRate;
  const hopSamples = Math.max(1, Math.floor((hopMs / 1000) * sr));
  const frameSamples = Math.max(1, Math.floor((frameMs / 1000) * sr));
  const chData = getMonoChannel(buffer);
  const n = chData.length;

  /** @type {number[]} */
  const env = [];
  for (let start = 0; start + frameSamples <= n; start += hopSamples) {
    let sum = 0;
    for (let i = 0; i < frameSamples; i += 1) {
      const s = chData[start + i];
      sum += s * s;
    }
    const rms = Math.sqrt(sum / frameSamples);
    env.push(rms);
  }

  // 簡單 energy threshold：去掉極小能量（視為 silence）
  if (env.length) {
    const maxVal = Math.max(...env);
    const thr = maxVal * 0.05; // 5% 以下視為 0
    for (let i = 0; i < env.length; i += 1) {
      if (env[i] < thr) env[i] = 0;
    }
  }

  return { env, hopSec: hopMs / 1000 };
}

/**
 * 簡單 moving average 平滑
 * @param {number[]} arr
 * @param {number} [win]
 * @returns {number[]}
 */
function movingAverage(arr, win = 7) {
  if (!arr.length) return [];
  const half = Math.floor(win / 2);
  const out = [];
  for (let i = 0; i < arr.length; i += 1) {
    let sum = 0;
    let count = 0;
    for (let k = i - half; k <= i + half; k += 1) {
      if (k < 0 || k >= arr.length) continue;
      sum += arr[k];
      count += 1;
    }
    out.push(count > 0 ? sum / count : 0);
  }
  return out;
}

/**
 * 對陣列做 z-normalization（平均 0、標準差 1），避免能量絕對大小影響相關值
 * @param {number[]} arr
 * @returns {number[]}
 */
function zNormalize(arr) {
  if (!arr.length) return [];
  let sum = 0;
  for (let i = 0; i < arr.length; i += 1) sum += arr[i];
  const mean = sum / arr.length;
  let varSum = 0;
  for (let i = 0; i < arr.length; i += 1) {
    const d = arr[i] - mean;
    varSum += d * d;
  }
  const std = Math.sqrt(varSum / Math.max(arr.length - 1, 1));
  if (std === 0) {
    return arr.map(() => 0);
  }
  return arr.map((v) => (v - mean) / std);
}

/**
 * 有限範圍 cross-correlation，回傳最佳 lag（frame 為單位）與 correlation 值
 * lag > 0 代表 target 需要往「晚」移動（target 晚進場）
 * @param {number[]} a master
 * @param {number[]} b target
 * @param {number} maxLagFrames
 * @returns {{ bestLag: number; bestCorr: number }}
 */
function crossCorrelateLimited(a, b, maxLagFrames) {
  const n = Math.min(a.length, b.length);
  if (!n || maxLagFrames <= 0) return { bestLag: 0, bestCorr: 0 };

  let bestLag = 0;
  let bestCorr = -Infinity;

  for (let lag = -maxLagFrames; lag <= maxLagFrames; lag += 1) {
    let sumAB = 0;
    let sumA2 = 0;
    let sumB2 = 0;
    let count = 0;

    if (lag >= 0) {
      for (let i = 0; i < n - lag; i += 1) {
        const va = a[i];
        const vb = b[i + lag];
        sumAB += va * vb;
        sumA2 += va * va;
        sumB2 += vb * vb;
        count += 1;
      }
    } else {
      const k = -lag;
      for (let i = 0; i < n - k; i += 1) {
        const va = a[i + k];
        const vb = b[i];
        sumAB += va * vb;
        sumA2 += va * va;
        sumB2 += vb * vb;
        count += 1;
      }
    }

    if (count === 0) continue;
    const denom = Math.sqrt(sumA2 * sumB2) || 1;
    const corr = sumAB / denom;
    if (corr > bestCorr) {
      bestCorr = corr;
      bestLag = lag;
    }
  }

  if (bestCorr === -Infinity) return { bestLag: 0, bestCorr: 0 };
  return { bestLag, bestCorr };
}

/**
 * RMS 自動對齊：以 energy envelope cross-correlation 對齊 target 相對 master
 * offsetSeconds > 0 代表 target 比 master 晚開始（target starts later）
 * @param {AudioBuffer} masterBuffer
 * @param {AudioBuffer} targetBuffer
 * @param {number} [maxShiftSec]
 * @returns {{ offsetSeconds: number; bestCorr: number }}
 */
function rmsAutoAlign(masterBuffer, targetBuffer, maxShiftSec = 3) {
  const hopMs = 20;
  const frameMs = 40;

  const masterEnv = computeRmsEnvelope(masterBuffer, hopMs, frameMs);
  const targetEnv = computeRmsEnvelope(targetBuffer, hopMs, frameMs);

  // 平滑後的 energy envelope
  let envA = movingAverage(masterEnv.env, 7);
  let envB = movingAverage(targetEnv.env, 7);

  const len = Math.min(envA.length, envB.length);
  if (!len) return { offsetSeconds: 0, bestCorr: 0 };
  envA = envA.slice(0, len);
  envB = envB.slice(0, len);

  // 標準 RMS correlation
  const a = zNormalize(envA);
  const b = zNormalize(envB);

  const hopSec = masterEnv.hopSec;
  const maxLagFrames = Math.max(1, Math.min(Math.floor(maxShiftSec / hopSec), len - 1));
  let { bestLag, bestCorr } = crossCorrelateLimited(a, b, maxLagFrames);

  // D) Fallback：在中度信心區間嘗試使用「onset-enhanced」能量差做額外 correlation
  if (bestCorr >= 0.15 && bestCorr <= 0.25) {
    /** onset-enhanced envelope: 只看上升部分 */
    const diffA = [];
    const diffB = [];
    for (let i = 0; i < len; i += 1) {
      if (i === 0) {
        diffA.push(0);
        diffB.push(0);
      } else {
        diffA.push(Math.max(0, envA[i] - envA[i - 1]));
        diffB.push(Math.max(0, envB[i] - envB[i - 1]));
      }
    }

    const diffASmoothed = movingAverage(diffA, 5);
    const diffBSmoothed = movingAverage(diffB, 5);
    const diffANorm = zNormalize(diffASmoothed);
    const diffBNorm = zNormalize(diffBSmoothed);

    const onsetCorr = crossCorrelateLimited(diffANorm, diffBNorm, maxLagFrames);

    // 若 onset-enhanced correlation 顯著更好，採用新的 offset
    if (onsetCorr.bestCorr > bestCorr + 0.05) {
      bestLag = onsetCorr.bestLag;
      bestCorr = onsetCorr.bestCorr;
    }
  }

  // 定義：offsetSeconds > 0 代表 target 晚於 master，因此要在播放時把 target 往前剪掉 offsetSeconds。
  const offsetSeconds = bestLag * hopSec;

  return { offsetSeconds, bestCorr };
}

/**
 * 偵測音檔的 Attack（開口）時間點與信心指標
 * 規則：
 * - 轉 mono
 * - frame 20ms / hop 10ms 的 RMS
 * - baseline = 前 0.5 秒 RMS 的 median
 * - mad = median(|env - baseline|)
 * - thr = max(baseline + mad * 6, peak * 0.15)
 * - 第一個「連續 3 frame > thr 且有上升趨勢」視為 onset
 * - conf = (peak - baseline) / (mad + 1e-6)
 * @param {AudioBuffer} buffer
 * @returns {{ timeSec: number; thr: number; baseline: number; mad: number; conf: number }}
 */
function detectAttackTime(buffer) {
  const sr = buffer.sampleRate;
  const frameMs = 20;
  const hopMs = 10;
  const frameSamples = Math.max(1, Math.floor((frameMs / 1000) * sr));
  const hopSamples = Math.max(1, Math.floor((hopMs / 1000) * sr));
  const chData = getMonoChannel(buffer);
  const n = chData.length;

  /** @type {number[]} */
  const env = [];
  for (let start = 0; start + frameSamples <= n; start += hopSamples) {
    let sum = 0;
    for (let i = 0; i < frameSamples; i += 1) {
      const s = chData[start + i];
      sum += s * s;
    }
    const rms = Math.sqrt(sum / frameSamples);
    env.push(rms);
  }

  if (!env.length) {
    return { timeSec: 0, thr: 0, baseline: 0, mad: 0, conf: 0 };
  }

  const hopSec = hopMs / 1000;
  const IGNORE_FIRST_SEC = 0.3;

  // smoothing
  const smoothEnv = movingAverage(env, 7);

  // baseline: 前 0.5 秒的 median
  const baselineFrames = Math.max(1, Math.floor(0.5 / hopSec));
  const early = smoothEnv.slice(0, Math.min(baselineFrames, smoothEnv.length));
  const sortedEarly = [...early].sort((x, y) => x - y);
  const medianEarly =
    sortedEarly.length % 2 === 1
      ? sortedEarly[(sortedEarly.length - 1) / 2]
      : (sortedEarly[sortedEarly.length / 2 - 1] + sortedEarly[sortedEarly.length / 2]) / 2;
  const baseline = medianEarly || 0;

  // mad: median(|env - baseline|)
  const absDev = smoothEnv.map((v) => Math.abs(v - baseline));
  const sortedDev = [...absDev].sort((x, y) => x - y);
  const mad =
    sortedDev.length % 2 === 1
      ? sortedDev[(sortedDev.length - 1) / 2]
      : (sortedDev[sortedDev.length / 2 - 1] + sortedDev[sortedDev.length / 2]) / 2;

  const peak = Math.max(...smoothEnv);
  const thr = Math.max(baseline + mad * 6, peak * 0.15);

  // onset: 第一個連續 N=3 frame > thr，且有上升趨勢
  const N = 3;
  let onsetIndex = 0;
  const ignoreFrames = Math.max(1, Math.floor(IGNORE_FIRST_SEC / hopSec));
  for (let i = ignoreFrames; i < smoothEnv.length - (N - 1); i += 1) {
    if (smoothEnv[i] <= thr) continue;
    let allAbove = true;
    for (let k = 0; k < N; k += 1) {
      if (smoothEnv[i + k] <= thr) {
        allAbove = false;
        break;
      }
    }
    if (!allAbove) continue;
    const slope = smoothEnv[i] - smoothEnv[i - 1];
    if (slope <= mad * 2) continue;
    onsetIndex = i;
    break;
  }

  const timeSec = onsetIndex * hopSec;
  const conf = (peak - baseline) / (mad + 1e-6);

  return { timeSec, thr, baseline, mad, conf };
}

/**
 * Attack Alignment：比較 master 與 user 的 attack 差異
 * 定義：attackOffsetSec = userAttack.timeSec - refAttack.timeSec
 * attackOffsetSec > 0 → USER 晚開口 → USER 需提前
 * @param {AudioBuffer} masterBuffer
 * @param {AudioBuffer} userBuffer
 * @returns {{ refAttack: { timeSec: number; thr: number; baseline: number; mad: number; conf: number }; userAttack: { timeSec: number; thr: number; baseline: number; mad: number; conf: number }; attackOffsetSec: number }}
 */
function attackAutoAlign(masterBuffer, userBuffer) {
  const refAttack = detectAttackTime(masterBuffer);
  const userAttack = detectAttackTime(userBuffer);
  const attackOffsetSec = userAttack.timeSec - refAttack.timeSec;
  return { refAttack, userAttack, attackOffsetSec };
}
/**
 * 允許的最大時間對齊位移（秒）
 * 例如 1.5 表示最多嘗試前後 1.5 秒的對齊
 * @type {number}
 */
const MAX_ALIGN_LAG_SECONDS = 1.5;

/**
 * 分段評分的段數
 * @type {number}
 */
const SEGMENT_COUNT = 4;

/**
 * 將頻率轉為 cent 基準（相對 440Hz，只是為了計算差值）
 * @param {number} freq
 * @returns {number}
 */
function freqToCents(freq) {
  if (!freq || freq <= 0) return NaN;
  return 1200 * Math.log2(freq / 440);
}

/**
 * 比較兩個頻率，在 cent 空間上的差值絕對值
 * @param {number | null} f1
 * @param {number | null} f2
 * @returns {number | null}
 */
function centDiff(f1, f2) {
  if (!f1 || !f2 || f1 <= 0 || f2 <= 0) return null;
  const c1 = freqToCents(f1);
  const c2 = freqToCents(f2);
  if (Number.isNaN(c1) || Number.isNaN(c2)) return null;
  return Math.abs(c1 - c2);
}

/**
 * 使用者相對標準音的偏差（有號），單位 cent。正=偏高，負=偏低
 * @param {number} refFreq
 * @param {number} userFreq
 * @returns {number | null}
 */
function centDeviationSigned(refFreq, userFreq) {
  if (!refFreq || !userFreq || refFreq <= 0 || userFreq <= 0) return null;
  return 1200 * Math.log2(userFreq / refFreq);
}

/** 音準偏差圖：顏色區間（|deviationCents|）≤25 綠、25–50 黃、>50 紅 */
const DEVIATION_GREEN = 25;
const DEVIATION_YELLOW = 50;

/**
 * 對音高序列做 moving average 平滑；無聲（≤minFreq 或 0）不參與平均，且維持為 0 以斷線
 * @param {number[]} pitches
 * @param {number} [windowSize]
 * @param {number} [minFreq]
 * @returns {number[]}
 */
function smoothPitches(pitches, windowSize = 5, minFreq = 60) {
  if (!pitches.length) return [];
  const half = Math.floor(windowSize / 2);
  /** @type {number[]} */
  const out = [];
  for (let i = 0; i < pitches.length; i += 1) {
    let sum = 0;
    let count = 0;
    for (let k = Math.max(0, i - half); k <= Math.min(pitches.length - 1, i + half); k += 1) {
      const v = pitches[k];
      if (v > minFreq) {
        sum += v;
        count += 1;
      }
    }
    out.push(count > 0 ? sum / count : 0);
  }
  return out;
}

/**
 * 對偏差序列做 moving average；null（無聲/無效）不參與平均且維持 null 以斷線
 * @param {(number | null)[]} deviations
 * @param {number} [windowSize]
 * @returns {(number | null)[]}
 */
function smoothDeviations(deviations, windowSize = 5) {
  if (!deviations.length) return [];
  const half = Math.floor(windowSize / 2);
  /** @type {(number | null)[]} */
  const out = [];
  for (let i = 0; i < deviations.length; i += 1) {
    let sum = 0;
    let count = 0;
    for (let k = Math.max(0, i - half); k <= Math.min(deviations.length - 1, i + half); k += 1) {
      const v = deviations[k];
      if (v != null && Number.isFinite(v)) {
        sum += v;
        count += 1;
      }
    }
    out.push(count > 0 ? sum / count : null);
  }
  return out;
}

/**
 * 從 File 讀成 ArrayBuffer
 * @param {File} file
 * @returns {Promise<ArrayBuffer>}
 */
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (reader.result instanceof ArrayBuffer) {
        resolve(reader.result);
      } else {
        reject(new Error("無法讀取檔案成 ArrayBuffer"));
      }
    };
    reader.onerror = () => reject(reader.error || new Error("讀取檔案失敗"));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 使用 AudioContext.decodeAudioData 解碼音檔
 * 同時支援 promise 版與 callback 版，並把錯誤詳細 log 出來
 * @param {ArrayBuffer} buffer
 * @returns {Promise<AudioBuffer>}
 */
async function decodeAudio(buffer) {
    const cloned = buffer.slice(0);
  
    try {
      // 新版瀏覽器：decodeAudioData 傳回 Promise（長度 = 1 個參數）
      if (audioContext.decodeAudioData.length === 1) {
        return await audioContext.decodeAudioData(cloned);
      }
  
      // 舊版：使用 callback 形式
      return await new Promise((resolve, reject) => {
        audioContext.decodeAudioData(
          cloned,
          (decoded) => resolve(decoded),
          (err) => reject(err || new Error("解碼音檔失敗"))
        );
      });
    } catch (e) {
      console.error("decodeAudio 發生錯誤", e);
      throw e;
    }
  }

/**
 * 將 AudioBuffer 轉為某一聲道的 Float32Array
 * 這裡只取第 0 聲道（左聲道）
 * @param {AudioBuffer} audioBuffer
 * @returns {Float32Array}
 */
function getMonoChannel(audioBuffer) {
  return audioBuffer.getChannelData(0);
}

/**
 * 使用簡單自相關法（autocorrelation）從一段音訊估計基頻
 * @param {Float32Array} frame
 * @param {number} sampleRate
 * @returns {number} 頻率（Hz），偵測不到時回傳 0
 */
function detectPitchAutoCorrelation(frame, sampleRate) {
  const size = frame.length;
  if (size < 32) return 0;

  // 計算音量，太小就當作無音高
  let sumSquares = 0;
  for (let i = 0; i < size; i += 1) {
    const v = frame[i];
    sumSquares += v * v;
  }
  const rms = Math.sqrt(sumSquares / size);
  if (rms < 0.01) return 0;

  // 只在合理的音高範圍內找：50Hz ~ 1000Hz
  const maxLag = Math.floor(sampleRate / 50);
  const minLag = Math.floor(sampleRate / 1000);
  if (maxLag >= size) return 0;

  let bestLag = -1;
  let bestVal = Infinity;

  for (let lag = minLag; lag <= maxLag; lag += 1) {
    let sum = 0;
    for (let i = 0; i < size - lag; i += 1) {
      const diff = frame[i] - frame[i + lag];
      sum += diff * diff;
    }
    if (sum < bestVal) {
      bestVal = sum;
      bestLag = lag;
    }
  }

  if (bestLag <= 0) return 0;
  return sampleRate / bestLag;
}

/**
 * 對 user 音高序列做時間對齊，找出最佳位移
 * 使用簡單 cross-correlation 概念，只在合理範圍內搜尋
 * @param {number[]} refPitches
 * @param {number[]} userPitches
 * @returns {{ ref: number[]; user: number[]; refOffsetSec: number; userOffsetSec: number }} 對齊後的兩個序列與播放起點偏移（秒）
 */
function alignPitchSeries(refPitches, userPitches) {
  const lenRef = refPitches.length;
  const lenUser = userPitches.length;
  const noOffset = { refOffsetSec: 0, userOffsetSec: 0 };
  if (lenRef === 0 || lenUser === 0) {
    return { ref: refPitches, user: userPitches, ...noOffset };
  }

  const frameDuration = SAMPLE_INTERVAL_SECONDS;
  const maxLagFrames = Math.min(
    Math.floor(MAX_ALIGN_LAG_SECONDS / frameDuration),
    Math.max(lenRef, lenUser) - 1
  );

  let bestLag = 0;
  let bestScore = -Infinity;

  // lag > 0: user 提前；lag < 0: user 落後
  for (let lag = -maxLagFrames; lag <= maxLagFrames; lag += 1) {
    let sumDiff = 0;
    let validCount = 0;
    let matchedCount = 0;

    const startRef = Math.max(0, lag);
    const startUser = Math.max(0, -lag);
    const localLen = Math.min(lenRef - startRef, lenUser - startUser);
    if (localLen <= 0) continue;

    for (let i = 0; i < localLen; i += 1) {
      const fRef = refPitches[startRef + i];
      const fUser = userPitches[startUser + i];
      if (!fRef || !fUser) continue;
      const d = centDiff(fRef, fUser);
      if (d == null) continue;
      validCount += 1;
      sumDiff += d;
      if (d <= CENT_TOLERANCE) matchedCount += 1;
    }

    if (validCount === 0) continue;

    const overlapRatio = validCount / localLen;
    if (overlapRatio < 0.3) continue;

    const avgDiff = sumDiff / validCount;
    const score = matchedCount - avgDiff / CENT_TOLERANCE;
    if (score > bestScore) {
      bestScore = score;
      bestLag = lag;
    }
  }

  if (bestScore === -Infinity || bestLag === 0) {
    const minLen = Math.min(lenRef, lenUser);
    return {
      ref: refPitches.slice(0, minLen),
      user: userPitches.slice(0, minLen),
      ...noOffset,
    };
  }

  if (bestLag > 0) {
    // user 提前：ref 從 bestLag 開始，播放時 ref 要從 buffer 的 bestLag*interval 秒處開始
    const trimmedRef = refPitches.slice(bestLag);
    const minLen = Math.min(trimmedRef.length, lenUser);
    return {
      ref: trimmedRef.slice(0, minLen),
      user: userPitches.slice(0, minLen),
      refOffsetSec: bestLag * SAMPLE_INTERVAL_SECONDS,
      userOffsetSec: 0,
    };
  }

  // bestLag < 0，user 落後：user 從 -bestLag 開始
  const trimmedUser = userPitches.slice(-bestLag);
  const minLen = Math.min(lenRef, trimmedUser.length);
  return {
    ref: refPitches.slice(0, minLen),
    user: trimmedUser.slice(0, minLen),
    refOffsetSec: 0,
    userOffsetSec: -bestLag * SAMPLE_INTERVAL_SECONDS,
  };
}

/**
 * 對一個 AudioBuffer 做多點音高偵測，傳回一串頻率陣列
 * @param {AudioBuffer} audioBuffer
 * @returns {number[]} 每一個元素是對應時間點的頻率（Hz），或 0 代表無法偵測
 */
function extractPitchSeries(audioBuffer) {
  const channelData = getMonoChannel(audioBuffer);
  const sampleRate = audioBuffer.sampleRate;

  // 每一段的樣本長度
  const frameSize = Math.floor(SAMPLE_INTERVAL_SECONDS * sampleRate);

  /** @type {number[]} */
  const pitches = [];

  for (let offset = 0; offset + frameSize <= channelData.length; offset += frameSize) {
    const frame = channelData.subarray(offset, offset + frameSize);
    const freq = detectPitchAutoCorrelation(frame, sampleRate) || 0;
    pitches.push(freq);
  }

  return pitches;
}

/**
 * 比對兩串音高序列，計算「在允許誤差內的比例」
 * @param {number[]} refPitches
 * @param {number[]} userPitches
 * @returns {number} 準確度（0 ~ 1）
 */
function computeAccuracy(refPitches, userPitches) {
  const len = Math.min(refPitches.length, userPitches.length);
  if (len === 0) return 0;

  let validCount = 0;
  let matchedCount = 0;

  for (let i = 0; i < len; i += 1) {
    const fRef = refPitches[i];
    const fUser = userPitches[i];

    // 若某一方完全偵測不到音高，直接跳過該 frame
    if (!fRef || !fUser) continue;

    const diff = centDiff(fRef, fUser);
    if (diff == null) continue;

    validCount += 1;
    if (diff <= CENT_TOLERANCE) {
      matchedCount += 1;
    }
  }

  if (validCount === 0) return 0;
  return matchedCount / validCount;
}

/**
 * 將對齊後的音高序列切成多段並計算每段準確度
 * @param {number[]} refPitches
 * @param {number[]} userPitches
 * @param {number} segments
 * @returns {number[]} 每段 0~1
 */
function computeSegmentAccuracies(refPitches, userPitches, segments) {
  const len = Math.min(refPitches.length, userPitches.length);
  if (len === 0 || segments <= 0) return [];

  const segLen = Math.floor(len / segments) || len;
  /** @type {number[]} */
  const result = [];

  for (let s = 0; s < segments; s += 1) {
    const start = s * segLen;
    const end = s === segments - 1 ? len : Math.min(len, start + segLen);
    if (start >= end) {
      result.push(0);
      continue;
    }
    const acc = computeAccuracy(
      refPitches.slice(start, end),
      userPitches.slice(start, end)
    );
    result.push(acc);
  }

  return result;
}

/**
 * 更新分段評分 UI
 * @param {number[]} segmentAccuracies 0~1
 */
function updateSegmentScores(segmentAccuracies) {
  if (!segmentScoresEl) return;
  if (!segmentAccuracies.length) {
    segmentScoresEl.textContent = "";
    return;
  }

  const parts = segmentAccuracies.map((acc, idx) => {
    const percentage = Math.round(acc * 100);
    let levelClass = "seg-bad";
    if (percentage >= 90) levelClass = "seg-good";
    else if (percentage >= 50) levelClass = "seg-mid";

    return `<div class="segment-pill ${levelClass}">
      <span class="seg-label">${idx + 1}</span>
      <span class="seg-value">${percentage}%</span>
    </div>`;
  });

  segmentScoresEl.innerHTML = parts.join("");
}

/**
 * 根據準確度更新文字（燈號已移除）
 * @param {number} accuracy 0~1
 */
function updateUIWithAccuracy(accuracy) {
  const percentage = Math.round(accuracy * 100);
  if (accuracyTextEl) {
    accuracyTextEl.innerHTML = `<span class="value">${percentage}%</span> 準確度`;
  }
}

/**
 * 秒數轉 mm:ss
 * @param {number} sec
 * @returns {string}
 */
function formatTime(sec) {
  const m = Math.floor(sec / 60);
  const s = Math.floor(sec % 60);
  return `${m}:${s < 10 ? "0" : ""}${s}`;
}

/** 自然音名 C D E F G A B 對應的 MIDI 餘數 */
const NATURAL_MIDI = [0, 2, 4, 5, 7, 9, 11];
const NOTE_NAMES = ["C", "D", "E", "F", "G", "A", "B"];

// 半音階音名（含升號），供 cents 視覺與 tooltip 使用
const CHROMATIC_NOTE_NAMES = [
  "C",
  "C#",
  "D",
  "D#",
  "E",
  "F",
  "F#",
  "G",
  "G#",
  "A",
  "A#",
  "B",
];

/**
 * 頻率(Hz) 轉 MIDI 音高（A4 = 440Hz = 69）
 * @param {number} freq
 * @returns {number}
 */
function freqToMidi(freq) {
  if (!freq || freq <= 0) return 0;
  return 69 + 12 * Math.log2(freq / 440);
}

/**
 * MIDI 音高轉自然音名＋八度（如 C4, A3）
 * @param {number} midi
 * @returns {string}
 */
function midiToNoteLabel(midi) {
  const m = Math.round(midi);
  const mod = m % 12;
  const idx = NATURAL_MIDI.indexOf(mod);
  if (idx === -1) return "";
  const octave = Math.floor(m / 12) - 1;
  return NOTE_NAMES[idx] + octave;
}

/**
 * MIDI 音高轉對應頻率(Hz)
 * @param {number} midi
 * @returns {number}
 */
function midiToFreq(midi) {
  return 440 * Math.pow(2, (midi - 69) / 12);
}

/**
 * 頻率(Hz) 轉 MIDI（wrapper，名稱符合顯示邏輯）
 * @param {number} hz
 * @returns {number}
 */
function hzToMidi(hz) {
  if (!hz || !Number.isFinite(hz) || hz <= 0) return NaN;
  return 69 + 12 * (Math.log(hz / 440) / Math.log(2));
}

/**
 * MIDI 轉頻率(Hz)（wrapper）
 * @param {number} midi
 * @returns {number}
 */
function midiToHz(midi) {
  if (!Number.isFinite(midi)) return NaN;
  return 440 * Math.pow(2, (midi - 69) / 12);
}

/**
 * 兩個頻率間的音分差（centsBetween(user, ref)）－顯示用
 * @param {number} hz
 * @param {number} refHz
 * @returns {number}
 */
function centsBetween(hz, refHz) {
  if (
    !Number.isFinite(hz) ||
    !Number.isFinite(refHz) ||
    hz <= 0 ||
    refHz <= 0
  ) {
    return 0;
  }
  return 1200 * Math.log2(hz / refHz);
}

/**
 * 將 octave error 折回同一八度範圍（僅供視覺化使用）
 * @param {number} cents
 * @returns {number}
 */
function wrapCentsToOctave(cents) {
  if (!Number.isFinite(cents)) return 0;
  // 折回 [-600, +600]，避免 1200c, -1200c 這種倍頻誤差爆圖
  while (cents > 600) cents -= 1200;
  while (cents < -600) cents += 1200;
  return cents;
}

// === 顯示層專用：統計/半音偏差/穩定化 helpers ===

/** 取中位數（給 stabilizeMidiTrack 用；避免缺 median 導致整張圖不畫） */
function median(arr) {
  const xs = arr
    .filter((v) => v != null && Number.isFinite(v))
    .slice()
    .sort((a, b) => a - b);
  if (!xs.length) return null;
  return xs[Math.floor(xs.length / 2)];
}

/**
 * 半音差值 wrap 到 [-6, +6]，避免跨八度顯示很醜
 * @param {number} deltaSemitones
 */
function wrapSemitoneDelta(deltaSemitones) {
  if (!Number.isFinite(deltaSemitones)) return deltaSemitones;
  let d = deltaSemitones;
  while (d > 6) d -= 12;
  while (d < -6) d += 12;
  return d;
}

/**
 * 以 ±50 cents 規則量化半音差：round 到最近半音（1 半音 = 100 cents）
 * @param {number} deltaSemitones
 */
function quantizeSemitoneBy50c(deltaSemitones) {
  if (!Number.isFinite(deltaSemitones)) return 0;
  return Math.round(deltaSemitones);
}

/**
 * voiced gate：把無聲/無效 pitch 變成 null，讓線段斷開
 * 目前先用 pitchHz 本身 gate；未來可接 RMS
 * @param {Array<number|null>} pitchHzArr
 * @param {number[]|null} rmsArr
 * @returns {Array<number|null>}
 */
function applyVoicedGate(pitchHzArr, rmsArr) {
  const out = new Array(pitchHzArr.length).fill(null);
  const RMS_THRESHOLD = 0.01;
  for (let i = 0; i < pitchHzArr.length; i += 1) {
    const hz = pitchHzArr[i];
    if (!Number.isFinite(hz) || hz <= 0) {
      out[i] = null;
      continue;
    }
    if (Array.isArray(rmsArr) && rmsArr.length === pitchHzArr.length) {
      const r = rmsArr[i];
      if (!Number.isFinite(r) || r < RMS_THRESHOLD) {
        out[i] = null;
        continue;
      }
    }
    out[i] = hz;
  }
  return out;
}

/**
 * 穩定 MIDI 軌：避免八度/諧波大跳
 * - octave 修正：選最接近上一幀的 (m, m±12, m±24)
 * - 大跳 gate：> 7 半音視為錯誤 => null（斷線）
 * - 5 點 median 平滑
 * @param {Array<number|null>} midiArr
 * @param {Array<number|null>=} refMidiArr
 * @returns {Array<number|null>}
 */
function stabilizeMidiTrack(midiArr, refMidiArr) {
  const n = midiArr.length;
  const out = new Array(n).fill(null);
  let prev = null;

  for (let i = 0; i < n; i += 1) {
    const m = midiArr[i];
    if (!Number.isFinite(m)) {
      out[i] = null;
      continue;
    }

    let candidate = m;

    if (prev != null) {
      const cands = [m, m + 12, m - 12, m + 24, m - 24];
      let best = cands[0];
      let bestDist = Math.abs(cands[0] - prev);
      for (const c of cands) {
        const d = Math.abs(c - prev);
        if (d < bestDist) {
          best = c;
          bestDist = d;
        }
      }
      candidate = best;

      // 離上一幀太遠 => 視為偵測錯誤
      if (bestDist > 7) {
        out[i] = null;
        continue;
      }
    }

    // 若提供 refMidiArr，限制 user 不能離 ref 太遠（避免顯示「明明同音卻亂跳」）
    if (Array.isArray(refMidiArr) && Number.isFinite(refMidiArr[i])) {
      const refM = refMidiArr[i];
      const dRef = Math.abs(candidate - refM);
      if (dRef > 7) {
        out[i] = null;
        continue;
      }
    }

    out[i] = candidate;
    prev = candidate;
  }

  // 5 點 median 平滑
  const medOut = new Array(n).fill(null);
  const WIN = 5;
  const HALF = Math.floor(WIN / 2);
  for (let i = 0; i < n; i += 1) {
    const win = [];
    for (let j = i - HALF; j <= i + HALF; j += 1) {
      if (j < 0 || j >= n) continue;
      const v = out[j];
      if (Number.isFinite(v)) win.push(v);
    }
    medOut[i] = win.length ? median(win) : null;
  }
  return medOut;
}

/**
 * 依 shiftSec 做 frame shift（只用於顯示，不改原陣列）
 * shiftSec > 0：往右移（前面補 null）
 * shiftSec < 0：往左移（後面補 null）
 * @param {Array<number|null>} series
 * @param {number} shiftSec
 * @param {number} sampleIntervalSec
 * @returns {Array<number|null>}
 */
function shiftSeriesBySec(series, shiftSec, sampleIntervalSec) {
  const n = series.length;
  const out = new Array(n).fill(null);
  if (!Number.isFinite(shiftSec) || !Number.isFinite(sampleIntervalSec) || sampleIntervalSec <= 0) {
    for (let i = 0; i < n; i += 1) out[i] = series[i];
    return out;
  }
  const shiftFrames = Math.round(shiftSec / sampleIntervalSec);
  for (let i = 0; i < n; i += 1) {
    const j = i - shiftFrames;
    out[i] = j >= 0 && j < n ? series[j] : null;
  }
  return out;
}

/**
 * MIDI 整數音高轉半音階音名（帶八度，如 C4, G#3）
 * @param {number} midi
 * @returns {string}
 */
function midiToNoteName(midi) {
  if (!Number.isFinite(midi)) return "";
  const m = Math.round(midi);
  const name = CHROMATIC_NOTE_NAMES[(m % 12 + 12) % 12];
  const octave = Math.floor(m / 12) - 1;
  return `${name}${octave}`;
}

/**
 * 依偏差絕對值回傳 stroke 顏色（綠/黃/紅）
 * @param {number} absCents
 * @returns {string}
 */
function deviationColor(absCents) {
  if (absCents <= DEVIATION_GREEN) return "rgba(34,197,94,0.95)";   // 綠
  if (absCents <= DEVIATION_YELLOW) return "rgba(234,179,8,0.95)";  // 黃
  return "rgba(239,68,68,0.95)";                                    // 紅
}

/**
 * 音準偏差圖（Deviation View）：標準線 y=0，使用者曲線為相對標準的偏差(cent)
 * @param {number[]} refPitches
 * @param {number[]} userPitches
 * @param {number} [playheadProgress] 0~1
 * @param {number} [durationSec] 總時長（供 tooltip 顯示時間）
 */
function drawPitchCurves(refPitches, userPitches, playheadProgress, durationSec) {
  if (!pitchCanvas) return;
  const ctx = pitchCanvas.getContext("2d");
  if (!ctx) return;

  const dpr = window.devicePixelRatio || 1;
  const cssWidth = Math.max(100, pitchCanvas.clientWidth || 380);
  const cssHeight = Math.max(80, pitchCanvas.clientHeight || 200);

  pitchCanvas.width = cssWidth * dpr;
  pitchCanvas.height = cssHeight * dpr;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

  ctx.clearRect(0, 0, cssWidth, cssHeight);

  // 先對 ref/user 做平滑，再決定實際長度，避免 index 超界
  const refS = smoothPitches(refPitches);
  const userS = smoothPitches(userPitches);
  const len = Math.min(refS.length, userS.length);
  if (!len) {
    lastPitchDrawData = null;
    return;
  }

  /** @type {(number | null)[]} 無聲/無效處為 null，斷線不畫到底 */
  let rawDeviations = [];
  let validCount = 0;

  // 第一輪：嚴格 voiced mask，使用平滑後的 refS/userS
  for (let i = 0; i < len; i += 1) {
    const refHz = refS[i];
    const userHz = userS[i];

    if (
      !Number.isFinite(refHz) ||
      refHz <= 0 ||
      !Number.isFinite(userHz) ||
      userHz <= 0
    ) {
      rawDeviations.push(null);
      continue;
    }

    let cents = centsBetween(userHz, refHz);
    if (!Number.isFinite(cents)) {
      rawDeviations.push(null);
      continue;
    }

    cents = wrapCentsToOctave(cents);
    rawDeviations.push(cents);
    validCount += 1;
  }

  // 若有效點過少，fallback：仍用平滑 Hz 算 cents，只在非 finite 時設為 null
  if (validCount < 5) {
    console.warn("Too few voiced points; fallback to non-masked deviations");
    rawDeviations = [];
    for (let i = 0; i < len; i += 1) {
      const refHz = refS[i];
      const userHz = userS[i];
      if (!Number.isFinite(refHz) || !Number.isFinite(userHz)) {
        rawDeviations.push(null);
        continue;
      }
      let cents = centsBetween(userHz, refHz);
      if (!Number.isFinite(cents)) {
        rawDeviations.push(null);
        continue;
      }
      cents = wrapCentsToOctave(cents);
      rawDeviations.push(cents);
    }
  }

  const deviations = smoothDeviations(rawDeviations);

  const paddingLeft = 36;
  const paddingRight = 8;
  const paddingTop = 12;
  const paddingBottom = 24;
  const w = cssWidth - paddingLeft - paddingRight;
  const h = cssHeight - paddingTop - paddingBottom;

  /** Y 軸範圍：±cents，0 在中央（至少涵蓋 ±50c 區域） */
  const maxCents = 200;
  const centerY = paddingTop + h / 2;

  /** 總時長（秒），x 軸 0..totalSec 對應整條曲目 */
  const totalSec = durationSec ?? len * SAMPLE_INTERVAL_SECONDS;

  /** x 軸依「時間」比例：第 i 格對應時間 i * SAMPLE_INTERVAL_SECONDS，支援完整曲目長度 */
  function xForIndex(i) {
    if (totalSec <= 0 || len <= 1) return paddingLeft;
    const timeAtI = i * SAMPLE_INTERVAL_SECONDS;
    return paddingLeft + (timeAtI / totalSec) * w;
  }

  /** deviationCents: 正=偏高、負=偏低；y 軸向上為正 */
  function yForCents(cents) {
    if (cents == null || !Number.isFinite(cents)) return centerY;
    const t = Math.max(-maxCents, Math.min(maxCents, cents)) / maxCents;
    return centerY - t * (h / 2);
  }

  lastPitchDrawData = {
    deviations,
    len,
    paddingLeft,
    w,
    durationSec: totalSec,
    ref: refS,
    user: userS,
  };

  // 背景
  ctx.fillStyle = "#020617";
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  // 三色誤差帶背景（以 y=0 為中心，±20/±50 cents）
  (function drawErrorBands() {
    const yGreenTop = yForCents(DEVIATION_GREEN);
    const yGreenBottom = yForCents(-DEVIATION_GREEN);
    const yYellowTop = yForCents(DEVIATION_YELLOW);
    const yYellowBottom = yForCents(-DEVIATION_YELLOW);
    const yMaxTop = yForCents(maxCents);
    const yMaxBottom = yForCents(-maxCents);

    // 綠帶：|cents| <= 20
    ctx.fillStyle = "rgba(34,197,94,0.10)";
    ctx.fillRect(
      paddingLeft,
      Math.min(yGreenTop, yGreenBottom),
      w,
      Math.abs(yGreenBottom - yGreenTop)
    );

    // 黃帶：20 < |cents| <= 50（上下兩側）
    ctx.fillStyle = "rgba(234,179,8,0.08)";
    // 上黃：20~50
    ctx.fillRect(
      paddingLeft,
      Math.min(yYellowTop, yGreenTop),
      w,
      Math.abs(yYellowTop - yGreenTop)
    );
    // 下黃：-20~-50
    ctx.fillRect(
      paddingLeft,
      Math.min(yGreenBottom, yYellowBottom),
      w,
      Math.abs(yGreenBottom - yYellowBottom)
    );

    // 紅帶：|cents| > 50（靠近上下邊緣）
    ctx.fillStyle = "rgba(239,68,68,0.06)";
    // 上紅：50~max
    ctx.fillRect(
      paddingLeft,
      Math.min(yMaxTop, yYellowTop),
      w,
      Math.abs(yMaxTop - yYellowTop)
    );
    // 下紅：-50~-max
    ctx.fillRect(
      paddingLeft,
      Math.min(yYellowBottom, yMaxBottom),
      w,
      Math.abs(yYellowBottom - yMaxBottom)
    );
  })();

  // 簡化水平輔助線：僅標示 ±20 / ±50 的邊界線（不顯示數字）
  ctx.strokeStyle = "rgba(15,23,42,0.8)";
  ctx.lineWidth = 1;
  const gridCents = [DEVIATION_YELLOW, DEVIATION_GREEN, -DEVIATION_GREEN, -DEVIATION_YELLOW];
  for (const c of gridCents) {
    const y = yForCents(c);
    ctx.beginPath();
    ctx.moveTo(paddingLeft, y);
    ctx.lineTo(cssWidth - paddingRight, y);
    ctx.stroke();
  }

  // 0 線：永遠顯示的基準線（偏高 / 偏低 的分界）
  const zeroY = yForCents(0);
  ctx.strokeStyle = "rgba(80,150,255,0.55)";
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(paddingLeft, zeroY);
  ctx.lineTo(cssWidth - paddingRight, zeroY);
  ctx.stroke();

  // 簡單文字提示：偏高 / 偏低（不顯示數字）
  ctx.fillStyle = "rgba(148,163,184,0.55)";
  ctx.font = "11px system-ui";
  ctx.textAlign = "left";
  ctx.textBaseline = "top";
  ctx.fillText("偏高", paddingLeft + 4, paddingTop + 2);
  ctx.textBaseline = "bottom";
  ctx.fillText("偏低", paddingLeft + 4, paddingTop + h - 2);

  // 使用者偏差曲線：依 |deviation| 變色，null 斷線（暫時先維持原樣，之後再調整三色誤差帶）
  ctx.lineWidth = 2;
  let prevX = 0;
  let prevY = 0;
  let prevColor = "";
  let started = false;
  for (let i = 0; i < len; i += 1) {
    const dev = deviations[i];
    if (dev == null || !Number.isFinite(dev)) {
      started = false;
      continue;
    }
    const absCents = Math.abs(dev);
    const color = deviationColor(absCents);
    const x = xForIndex(i);
    const y = yForCents(dev);

    if (!started) {
      ctx.strokeStyle = color;
      ctx.beginPath();
      ctx.moveTo(x, y);
      started = true;
      prevColor = color;
    } else {
      if (color !== prevColor) {
        ctx.stroke();
        ctx.strokeStyle = color;
        ctx.beginPath();
        ctx.moveTo(prevX, prevY);
        ctx.lineTo(x, y);
        prevColor = color;
      } else {
        ctx.lineTo(x, y);
      }
    }
    prevX = x;
    prevY = y;
  }
  ctx.stroke();

  // 圖例
  const legendY = cssHeight - 6;
  ctx.textAlign = "left";
  ctx.textBaseline = "bottom";
  ctx.font = "13px system-ui";
  ctx.fillStyle = "#38bdf8";
  ctx.fillRect(paddingLeft, legendY - 8, 14, 2);
  ctx.fillStyle = "#cbd5f5";
  ctx.fillText("標準 (In Tune)", paddingLeft + 20, legendY);

  ctx.fillStyle = "#22c55e";
  ctx.fillRect(paddingLeft + 110, legendY - 8, 12, 2);
  ctx.fillStyle = "#cbd5f5";
  ctx.fillText("≤20¢", paddingLeft + 126, legendY);
  ctx.fillStyle = "#eab308";
  ctx.fillRect(paddingLeft + 162, legendY - 8, 12, 2);
  ctx.fillStyle = "#cbd5f5";
  ctx.fillText("20–50¢", paddingLeft + 178, legendY);
  ctx.fillStyle = "#ef4444";
  ctx.fillRect(paddingLeft + 222, legendY - 8, 12, 2);
  ctx.fillStyle = "#cbd5f5";
  ctx.fillText(">50¢", paddingLeft + 238, legendY);

  // 播放線與已播區域
  if (playheadProgress !== undefined && playheadProgress >= 0 && playheadProgress <= 1) {
    const playheadX = paddingLeft + playheadProgress * w;
    ctx.fillStyle = "rgba(56, 189, 248, 0.1)";
    ctx.fillRect(paddingLeft, paddingTop, playheadX - paddingLeft, h);
    ctx.strokeStyle = "#38bdf8";
    ctx.lineWidth = 2;
    ctx.setLineDash([4, 4]);
    ctx.beginPath();
    ctx.moveTo(playheadX, paddingTop);
    ctx.lineTo(playheadX, paddingTop + h);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

/**
 * 清空圖像、分段、時間軸
 */
function clearVisualizations() {
  if (pitchCanvas) {
    const ctx = pitchCanvas.getContext("2d");
    if (ctx) {
      const cssWidth = pitchCanvas.clientWidth || 400;
      const cssHeight = pitchCanvas.clientHeight || 200;
      ctx.clearRect(0, 0, cssWidth, cssHeight);
    }
  }
  if (segmentScoresEl) segmentScoresEl.textContent = "";
  playbackOffsetSec = 0;
  if (timelineEl) timelineEl.value = "0";
  if (timeCurrentEl) timeCurrentEl.textContent = "0:00";
  if (timeTotalEl) timeTotalEl.textContent = "0:00";
}

/**
 * 主要流程：讀檔 → 解碼 → 抽取音高 → 對齊 → 比對 → 更新 UI
 */
async function handleAnalyze() {
  if (!refInput || !userInput || !analyzeBtn) return;

  const refFile = refInput.files && refInput.files[0];
  const userFile = userInput.files && userInput.files[0];

  if (!refFile) {
    alert("請先選擇標準音檔。");
    return;
  }

  if (!userFile && !recordedUserAudio) {
    alert("請上傳你的錄音檔，或先用麥克風錄一段。");
    return;
  }

  try {
    analyzeBtn.disabled = true;
    analyzeBtn.textContent = "分析中...";
    clearVisualizations();

    // 1. 讀取並解碼標準音
    const refBuf = await readFileAsArrayBuffer(refFile);
    const refAudio = await decodeAudio(refBuf);

    // 2. 準備使用者音訊：檔案優先，其次麥克風錄音
    /** @type {AudioBuffer} */
    let userAudio;
    if (userFile) {
      const userBuf = await readFileAsArrayBuffer(userFile);
      userAudio = await decodeAudio(userBuf);
    } else if (recordedUserAudio) {
      userAudio = recordedUserAudio;
    } else {
      throw new Error("沒有可用的使用者音訊。");
    }

    // 3. 抽取音高序列
    const refPitchesRaw = extractPitchSeries(refAudio);
    const userPitchesRaw = extractPitchSeries(userAudio);

    // 4. 先做時間對齊，再計算整體與分段準確度
    const aligned = alignPitchSeries(refPitchesRaw, userPitchesRaw);
    const accuracy = computeAccuracy(aligned.ref, aligned.user);
    const segmentAccuracies = computeSegmentAccuracies(
      aligned.ref,
      aligned.user,
      SEGMENT_COUNT
    );

    // 5. 更新 UI：準確度文字、曲線圖、分段分數、總時長（以實際音檔長度為準，支援完整曲目）
    const totalSec = Math.min(refAudio.duration, userAudio.duration);
    updateUIWithAccuracy(accuracy);
    drawPitchCurves(aligned.ref, aligned.user, undefined, totalSec);
    updateSegmentScores(segmentAccuracies);
    if (timeTotalEl) timeTotalEl.textContent = formatTime(totalSec);

    // 6. 存起來供同步播放與時間軸（含時間差供對齊播放）
    lastPlaybackData = {
      refBuffer: refAudio,
      userBuffer: userAudio,
      refPitches: aligned.ref,
      userPitches: aligned.user,
      durationSec: totalSec,
      refOffsetSec: aligned.refOffsetSec ?? 0,
      userOffsetSec: aligned.userOffsetSec ?? 0,
    };
    playbackOffsetSec = 0;
    if (timelineEl) timelineEl.value = "0";

    // 8. 分句偵測：優先使用 refPitches，若 ref 完全無效則改用 userPitches
    const hasValidRef = aligned.ref.some((p) => p && Number.isFinite(p) && p > 0);
    const phraseSource = hasValidRef ? aligned.ref : aligned.user;
    phrases = detectPhrases(phraseSource, totalSec);

    // 9. 分句條：根據最新曲目長度重繪 phrase bar
    renderPhraseBar();

    // 7. 合併功能：比對完自動開始播放並顯示動畫（含背景音樂，由 startSyncedPlayback 一併啟動）
    startSyncedPlayback();
    updateReplayButtonState();
  } catch (err) {
    console.error("分析時發生錯誤", err);
    const msg =
      err && typeof err === "object" && "message" in err
        ? /** @type {Error} */ (err).message
        : String(err);
    alert(`分析時發生錯誤：${msg}`);
  } finally {
    analyzeBtn.disabled = false;
    analyzeBtn.textContent = isPlaying ? "停止播放" : "開始比對並播放";
  }
}

if (analyzeBtn) {
  analyzeBtn.addEventListener("click", async () => {
    if (audioContext.state === "suspended") {
      await audioContext.resume().catch(() => {});
    }
    if (isPlaying) {
      stopPlayback();
    } else if (lastPlaybackData) {
      startSyncedPlayback();
    } else {
      await handleAnalyze();
    }
  });
}

/**
 * 停止同步播放並還原 UI（也會順便關掉 BG 軌）
 */
function stopPlayback() {
  if (!isPlaying && !isBgPlaying) return;

  if (lastPlaybackData) {
    const T = Math.max(0, audioContext.currentTime - masterStartTime);
    playbackOffsetSec = Math.min(T, lastPlaybackData.durationSec);
  }

  isPlaying = false;

  if (refSourceNode) {
    try { refSourceNode.stop(); refSourceNode.disconnect(); } catch (_) {}
    refSourceNode = null;
  }
  if (userSourceNode) {
    try { userSourceNode.stop(); userSourceNode.disconnect(); } catch (_) {}
    userSourceNode = null;
  }
  if (activeRefSource) {
    try { activeRefSource.stop(); activeRefSource.disconnect(); } catch (_) {}
    activeRefSource = null;
  }
  if (activeUserSource) {
    try { activeUserSource.stop(); activeUserSource.disconnect(); } catch (_) {}
    activeUserSource = null;
  }
  if (activeAccSource) {
    try { activeAccSource.stop(); activeAccSource.disconnect(); } catch (_) {}
    activeAccSource = null;
  }
  if (refGainNode) {
    try { refGainNode.disconnect(); } catch (_) {}
    refGainNode = null;
  }
  if (userGainNode) {
    try { userGainNode.disconnect(); } catch (_) {}
    userGainNode = null;
  }
  if (playbackAnimationId) {
    cancelAnimationFrame(playbackAnimationId);
    playbackAnimationId = 0;
  }
  stopBgMusic();
  if (analyzeBtn) analyzeBtn.textContent = "開始比對並播放";
  if (lastPlaybackData) {
    const pct = playbackOffsetSec / lastPlaybackData.durationSec;
    drawPitchCurves(
      lastPlaybackData.refPitches,
      lastPlaybackData.userPitches,
      pct,
      lastPlaybackData.durationSec
    );
  }
  if (timelineEl) timelineEl.value = String((playbackOffsetSec / (lastPlaybackData?.durationSec || 1)) * 100);
  if (timeCurrentEl && lastPlaybackData) timeCurrentEl.textContent = formatTime(playbackOffsetSec);
}

/**
 * 播放時每幀更新播放線與時間軸
 */
function animatePlayhead() {
  if (!isPlaying || !lastPlaybackData) return;
  const now = audioContext.currentTime;
  const currentSec = Math.max(0, now - masterStartTime);
  const clampedSec = Math.min(currentSec, lastPlaybackData.durationSec);
  const progress = Math.min(1, Math.max(0, clampedSec / lastPlaybackData.durationSec));

  // A/B Loop：若啟用且到達 B，觸發一次跳回 A 的流程
  if (
    isPlaying &&
    abLoop.enabled &&
    !abLoopJumping &&
    abLoop.aSec != null &&
    abLoop.bSec != null &&
    Number.isFinite(abLoop.aSec) &&
    Number.isFinite(abLoop.bSec) &&
    abLoop.aSec < abLoop.bSec &&
    clampedSec >= abLoop.bSec
  ) {
    triggerAbLoopJump();
    return;
  }

  drawPitchCurves(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    progress,
    lastPlaybackData.durationSec
  );

  updatingTimelineFromPlayback = true;
  if (timelineEl) timelineEl.value = String(progress * 100);
  if (timeCurrentEl) timeCurrentEl.textContent = formatTime(clampedSec);
  updatingTimelineFromPlayback = false;

  // 播放時高亮當前所在的分句區塊
  if (phraseBarEl && phrases.length && lastPlaybackData) {
    const segEls = phraseBarEl.querySelectorAll(".phrase-seg");
    let activeIndex = -1;
    for (let i = 0; i < phrases.length; i += 1) {
      const p = phrases[i];
      if (clampedSec >= p.startSec && clampedSec < p.endSec) {
        activeIndex = i;
        break;
      }
    }
    segEls.forEach((el, idx) => {
      if (idx === activeIndex) {
        el.classList.add("active");
      } else {
        el.classList.remove("active");
      }
    });
  }

  if (progress >= 1) {
    playbackOffsetSec = lastPlaybackData.durationSec;
    stopPlayback();
    return;
  }
  playbackAnimationId = requestAnimationFrame(animatePlayhead);
}

/**
 * Master Timeline：給定目前播放位置與 finalOffsetSec，計算 Task1 版 startAt/outputDelay。
 * 定義：
 *   T = playbackOffsetSec
 *   trackTimeNow = T - finalOffsetSec
 *
 * 非循環（ref/user）：
 *   trackTimeNow < 0 → startAtSec = 0；outputDelaySec = -trackTimeNow
 *   否則          → startAtSec = trackTimeNow；outputDelaySec = 0
 *
 * 循環（BG loop）：
 *   trackTimeNow < 0 → startAtSec = 0；outputDelaySec = -trackTimeNow
 *   否則          → startAtSec = ((trackTimeNow % dur) + dur) % dur；outputDelaySec = 0
 *
 * 排程統一使用：
 *   source.start(audioContext.currentTime + outputDelaySec, startAtSec)
 *
 * @param {number} playbackOffsetSec
 * @param {number} finalOffsetSec autoOffsetSec + manualOffsetSec
 * @param {number} bufferDurationSec
 * @param {boolean} loop
 * @returns {{ startAtSec: number; outputDelaySec: number }}
 */
function computeStartAtAndOutputTime(
  playbackOffsetSec,
  finalOffsetSec,
  bufferDurationSec,
  loop
) {
  const trackTimeNow = playbackOffsetSec - finalOffsetSec;

  if (!loop) {
    const startAtSec = Math.max(0, trackTimeNow);
    const outputDelaySec = Math.max(0, -trackTimeNow);
    return { startAtSec, outputDelaySec };
  }

  if (trackTimeNow < 0) {
    return { startAtSec: 0, outputDelaySec: -trackTimeNow };
  }

  const dur = Math.max(bufferDurationSec, 0.001);
  const startAtSec = ((trackTimeNow % dur) + dur) % dur;
  return { startAtSec, outputDelaySec: 0 };
}

/**
 * 同步播放 ref / user / acc：delay 只影響 source.start 時間，不依賴舊 offset。
 * @param {{ refBuffer: AudioBuffer; userBuffer: AudioBuffer; accBuffer?: AudioBuffer; durationSec: number } | undefined} [payload] 未傳則從 lastPlaybackData 建
 */
function startSyncedPlayback(payload) {
  if (!payload && lastPlaybackData) {
    payload = {
      refBuffer: lastPlaybackData.refBuffer,
      userBuffer: lastPlaybackData.userBuffer,
      accBuffer: masterPianoBuffer || undefined,
      durationSec: lastPlaybackData.durationSec,
    };
  }
  if (!payload || !payload.refBuffer || !payload.userBuffer) return;
  if (isPlaying) return;

  stopPlayback();

  if (audioContext.state === "suspended") {
    audioContext.resume().catch(() => {});
  }

  const { refBuffer, userBuffer, accBuffer, durationSec } = payload;
  const refDelay = parseFloat(refDelayInput?.value || 0) || 0;
  const userDelay = parseFloat(userDelayInput?.value || 0) || 0;

  console.log("PLAY WITH DELAY:", refDelay, userDelay);

  masterStartTime = audioContext.currentTime + 0.05;
  contentStartTime = masterStartTime;
  playbackOffsetSec = 0;

  const refVol = refVolumeInput ? parseInt(refVolumeInput.value, 10) / 100 : 1;
  const userVol = userVolumeInput ? parseInt(userVolumeInput.value, 10) / 100 : 1;

  refGainNode = audioContext.createGain();
  refGainNode.gain.value = refVol;
  refGainNode.connect(audioContext.destination);
  userGainNode = audioContext.createGain();
  userGainNode.gain.value = userVol;
  userGainNode.connect(audioContext.destination);

  if (refBuffer) {
    const refSource = audioContext.createBufferSource();
    refSource.buffer = refBuffer;
    refSource.connect(refGainNode);
    refSource.start(masterStartTime + refDelay, 0, durationSec);
    activeRefSource = refSource;
  }

  if (accBuffer) {
    const accSource = audioContext.createBufferSource();
    accSource.buffer = accBuffer;
    accSource.connect(audioContext.destination);
    accSource.start(masterStartTime, 0, durationSec);
    activeAccSource = accSource;
  }

  if (userBuffer) {
    const userSource = audioContext.createBufferSource();
    userSource.buffer = userBuffer;
    userSource.connect(userGainNode);
    userSource.start(masterStartTime + userDelay, 0, durationSec);
    activeUserSource = userSource;
  }

  isPlaying = true;
  if (analyzeBtn) analyzeBtn.textContent = "停止播放";
  if (bgBuffers.some(Boolean)) startBgMusic(0);
  playbackAnimationId = requestAnimationFrame(animatePlayhead);
}

/**
 * 依時間軸進度 seek：可拖曳選擇播放起點
 */
function seekToPercent(pct) {
  if (!lastPlaybackData) return;
  pct = Math.max(0, Math.min(1, pct));
  playbackOffsetSec = pct * lastPlaybackData.durationSec;
  if (timelineEl) timelineEl.value = String(pct * 100);
  if (timeCurrentEl) timeCurrentEl.textContent = formatTime(playbackOffsetSec);
  drawPitchCurves(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    pct,
    lastPlaybackData.durationSec
  );
}

if (timelineEl) {
  timelineEl.addEventListener("input", () => {
    if (updatingTimelineFromPlayback) return;
    if (isProRecording) return;
    const pct = parseFloat(timelineEl.value) / 100;
    const wasPlaying = isPlaying;
    if (wasPlaying) {
      stopPlayback();
    }
    seekToPercent(pct);
    if (wasPlaying) {
      startSyncedPlayback();
    }
  });
}

/**
 * 取得目前播放頭位置（秒）
 * - 播放中：T = audioContext.currentTime - masterStartTime
 * - 未播放：使用 playbackOffsetSec
 * 結果會 clamp 到 [0, durationSec]
 * @returns {number}
 */
function getCurrentPlayheadSec() {
  if (!lastPlaybackData) return 0;
  const dur = lastPlaybackData.durationSec || 0;
  if (dur <= 0) return 0;
  let t = 0;
  if (isPlaying) {
    t = Math.max(0, audioContext.currentTime - masterStartTime);
  } else {
    t = Math.max(0, playbackOffsetSec);
  }
  if (t > dur) t = dur;
  return t;
}

/**
 * 更新 A/B Loop label 顯示
 */
function updateAbLoopLabels() {
  if (abALabelEl) {
    abALabelEl.textContent =
      abLoop.aSec != null && Number.isFinite(abLoop.aSec)
        ? formatTime(abLoop.aSec)
        : "--:--";
  }
  if (abBLabelEl) {
    abBLabelEl.textContent =
      abLoop.bSec != null && Number.isFinite(abLoop.bSec)
        ? formatTime(abLoop.bSec)
        : "--:--";
  }
  if (abEnableCheckbox) {
    abEnableCheckbox.checked = !!abLoop.enabled;
  }
  if (abGapInput) {
    abGapInput.value = String(abLoop.gapSec.toFixed(2));
  }
}

// 初始化 A/B Loop UI 狀態
updateAbLoopLabels();

// A/B Loop 按鈕事件
if (abSetABtn) {
  abSetABtn.addEventListener("click", () => {
    if (isProRecording) {
      alert("錄音期間暫不支援 A/B Loop。");
      return;
    }
    if (!lastPlaybackData) return;
    const t = getCurrentPlayheadSec();
    abLoop.aSec = t;
    // 若已存在 B，檢查 A<B
    if (abLoop.bSec != null && abLoop.aSec >= abLoop.bSec) {
      alert("A 必須小於 B。請重新設定。");
      abLoop.enabled = false;
      if (abEnableCheckbox) abEnableCheckbox.checked = false;
    }
    updateAbLoopLabels();
  });
}

if (abSetBBtn) {
  abSetBBtn.addEventListener("click", () => {
    if (isProRecording) {
      alert("錄音期間暫不支援 A/B Loop。");
      return;
    }
    if (!lastPlaybackData) return;
    const t = getCurrentPlayheadSec();
    abLoop.bSec = t;
    if (abLoop.aSec != null && abLoop.aSec >= abLoop.bSec) {
      alert("A 必須小於 B。請重新設定。");
      abLoop.enabled = false;
      if (abEnableCheckbox) abEnableCheckbox.checked = false;
    }
    updateAbLoopLabels();
  });
}

if (abClearBtn) {
  abClearBtn.addEventListener("click", () => {
    if (isProRecording) {
      alert("錄音期間暫不支援 A/B Loop。");
      return;
    }
    abLoop.aSec = null;
    abLoop.bSec = null;
    abLoop.enabled = false;
    if (abEnableCheckbox) abEnableCheckbox.checked = false;
    updateAbLoopLabels();
  });
}

if (abEnableCheckbox) {
  abEnableCheckbox.addEventListener("change", () => {
    if (isProRecording) {
      abEnableCheckbox.checked = abLoop.enabled;
      alert("錄音期間暫不支援 A/B Loop。");
      return;
    }
    // 只有在 A < B 且兩者都存在時才允許啟用
    if (
      abLoop.aSec == null ||
      abLoop.bSec == null ||
      !Number.isFinite(abLoop.aSec) ||
      !Number.isFinite(abLoop.bSec) ||
      abLoop.aSec >= abLoop.bSec
    ) {
      abLoop.enabled = false;
      abEnableCheckbox.checked = false;
      alert("請先設定有效的 A 與 B（且 A < B）才能啟用 Loop。");
      return;
    }
    abLoop.enabled = abEnableCheckbox.checked;
  });
}

if (abGapInput) {
  abGapInput.addEventListener("change", () => {
    let v = parseFloat(abGapInput.value);
    if (!Number.isFinite(v)) v = 0.3;
    if (v < 0) v = 0;
    if (v > 1.0) v = 1.0;
    abLoop.gapSec = v;
    abGapInput.value = String(v.toFixed(2));
  });
}

// A/B Loop 快捷鍵（Shift+A / Shift+B / Shift+L）
document.addEventListener("keydown", (e) => {
  if (!e.shiftKey) return;
  if (e.repeat) return;
  if (e.target && (/** @type {HTMLElement} */ (e.target)).tagName === "INPUT") return;

  if (e.code === "KeyA") {
    if (abSetABtn) abSetABtn.click();
  } else if (e.code === "KeyB") {
    if (abSetBBtn) abSetBBtn.click();
  } else if (e.code === "KeyL") {
    if (isProRecording) {
      alert("錄音期間暫不支援 A/B Loop。");
      return;
    }
    if (!abEnableCheckbox) return;
    // 嘗試切換勾選；change handler 會驗證 A/B 是否有效
    abEnableCheckbox.checked = !abEnableCheckbox.checked;
    const event = new Event("change");
    abEnableCheckbox.dispatchEvent(event);
  }
});

/**
 * 分句拖曳 mousemove：根據滑鼠位移調整 phrases[i] 的 startSec / endSec
 */
document.addEventListener("mousemove", (e) => {
  if (phraseDragIndex < 0) return;
  if (!lastPlaybackData || !phraseBarEl) return;
  const duration = lastPlaybackData.durationSec || 0;
  if (!duration || duration <= 0) return;

  const dx = e.clientX - phraseDragStartX;
  const width = phraseDragTimelineWidth || phraseBarEl.getBoundingClientRect().width || 1;
  if (!width) return;

  const deltaSec = (dx / width) * duration;
  const i = phraseDragIndex;
  if (i < 0 || i >= phrases.length) return;

  const p = phrases[i];
  let newStart = phraseDragStartStartSec;
  let newEnd = phraseDragStartEndSec;

  if (phraseDragIsLeft) {
    newStart = phraseDragStartStartSec + deltaSec;
    const prevEnd = i > 0 ? phrases[i - 1].endSec : 0;
    const maxStart = newEnd - PHRASE_MIN_ADJUST_SEC;
    if (newStart < prevEnd) newStart = prevEnd;
    if (newStart > maxStart) newStart = maxStart;
    if (newStart < 0) newStart = 0;
    p.startSec = newStart;
  } else {
    newEnd = phraseDragStartEndSec + deltaSec;
    const nextStart = i < phrases.length - 1 ? phrases[i + 1].startSec : duration;
    const minEnd = newStart + PHRASE_MIN_ADJUST_SEC;
    if (newEnd > nextStart) newEnd = nextStart;
    if (newEnd < minEnd) newEnd = minEnd;
    if (newEnd > duration) newEnd = duration;
    p.endSec = newEnd;
  }

  renderPhraseBar();
});

/**
 * 分句拖曳 mouseup：清除拖曳狀態
 */
document.addEventListener("mouseup", () => {
  phraseDragIndex = -1;
});

/**
 * 觸發一次 A/B Loop 跳轉：
 * - 停止目前播放（stopPlayback）
 * - 將播放頭移到 A
 * - 等待 gapSec 秒後，再重新 startSyncedPlayback
 */
function triggerAbLoopJump() {
  if (!lastPlaybackData) return;
  if (abLoopJumping) return;
  if (
    !abLoop.enabled ||
    abLoop.aSec == null ||
    abLoop.bSec == null ||
    !Number.isFinite(abLoop.aSec) ||
    !Number.isFinite(abLoop.bSec) ||
    abLoop.aSec >= abLoop.bSec
  ) {
    return;
  }

  abLoopJumping = true;

  // 1) 停止目前播放，讓 stopPlayback 回寫 playbackOffsetSec
  if (isPlaying) {
    stopPlayback();
  }

  // 2) 將播放頭移到 A
  const dur = lastPlaybackData.durationSec || 0;
  let aSec = Math.max(0, Math.min(abLoop.aSec, dur));
  if (!Number.isFinite(aSec)) aSec = 0;
  const pct = dur > 0 ? aSec / dur : 0;
  seekToPercent(pct);

  // 3) 若使用者在等待期間關掉 abLoop 或 A/B 變得無效，就不要重啟
  const gapMs = Math.max(0, Math.min(abLoop.gapSec || 0, 1.0)) * 1000;

  const restart = () => {
    if (
      !abLoop.enabled ||
      abLoop.aSec == null ||
      abLoop.bSec == null ||
      !Number.isFinite(abLoop.aSec) ||
      !Number.isFinite(abLoop.bSec) ||
      abLoop.aSec >= abLoop.bSec
    ) {
      abLoopJumping = false;
      return;
    }
    startSyncedPlayback();
    abLoopJumping = false;
  };

  if (gapMs > 0) {
    setTimeout(restart, gapMs);
  } else {
    restart();
  }
}

/** Pitch 偏差圖 tooltip：時間、目標音/實際音、偏差(± cents)、半音趨勢提示 */
function setupPitchTooltip() {
  const wrap = document.getElementById("pitch-graph-wrap");
  if (!pitchCanvas || !pitchTooltipEl || !wrap) return;

  wrap.addEventListener("mousemove", (e) => {
    if (!lastPitchDrawData) {
      pitchTooltipEl.classList.remove("visible");
      return;
    }
    const wrapRect = wrap.getBoundingClientRect();
    const canvasRect = pitchCanvas.getBoundingClientRect();
    const x = e.clientX - canvasRect.left;
    const { paddingLeft, w, len, deviations, durationSec, ref, user } = lastPitchDrawData;
    if (len <= 1) {
      pitchTooltipEl.classList.remove("visible");
      return;
    }
    const t = Math.max(0, Math.min(1, (x - paddingLeft) / w));
    const timeSec = t * durationSec;
    const i = Math.floor(Math.max(0, Math.min(len - 1, timeSec / SAMPLE_INTERVAL_SECONDS)));
    const dev = deviations[i];
    const timeStr = formatTime(timeSec);
    if (dev == null || !Number.isFinite(dev)) {
      pitchTooltipEl.innerHTML = `${timeStr}<br><span class="tooltip-muted">—</span>`;
    } else {
      const sign = dev >= 0 ? "+" : "";
      const centsRounded = Math.round(dev);
      const dir = dev > 0 ? "偏高" : dev < 0 ? "偏低" : "唱準";

      let refHz = ref && Array.isArray(ref) ? ref[i] : null;
      let userHz = user && Array.isArray(user) ? user[i] : null;
      if (!refHz || refHz <= 0 || !Number.isFinite(refHz)) refHz = null;
      if (!userHz || userHz <= 0 || !Number.isFinite(userHz)) userHz = null;

      let refNote = "";
      let userNote = "";
      let crossText = "";

      if (refHz && userHz) {
        const refMidi = hzToMidi(refHz);
        const userMidi = hzToMidi(userHz);
        refNote = midiToNoteName(refMidi);
        userNote = midiToNoteName(userMidi);

        if (dev > DEVIATION_YELLOW) {
          const nextNote = midiToNoteName(refMidi + 1);
          crossText = `偏高，趨近 ${nextNote}`;
        } else if (dev < -DEVIATION_YELLOW) {
          const prevNote = midiToNoteName(refMidi - 1);
          crossText = `偏低，趨近 ${prevNote}`;
        } else {
          crossText = "落在目標音範圍內";
        }
      } else {
        crossText = dir;
      }

      const noteLine =
        refNote || userNote
          ? `目標: ${refNote || "—"} ｜ 你: ${userNote || "—"}`
          : "";

      pitchTooltipEl.innerHTML = `${timeStr}${
        noteLine ? "<br>" + noteLine : ""
      }<br>${sign}${centsRounded} cents<br>${crossText}`;
    }
    pitchTooltipEl.classList.add("visible");
    pitchTooltipEl.style.left = `${e.clientX - wrapRect.left + 12}px`;
    pitchTooltipEl.style.top = `${e.clientY - wrapRect.top + 8}px`;
  });

  wrap.addEventListener("mouseleave", () => {
    pitchTooltipEl.classList.remove("visible");
  });
}
setupPitchTooltip();

/**
 * 停止背景音樂
 */
function stopBgMusic() {
  if (!isBgPlaying) return;
  isBgPlaying = false;
  for (const node of bgSourceNodes) {
    try {
      node.stop();
      node.disconnect();
    } catch (_) {}
  }
  bgSourceNodes = [];
  for (const g of bgTrackGainNodes) {
    if (g) {
      try { g.disconnect(); } catch (_) {}
    }
  }
  bgTrackGainNodes = [];
  if (bgPlayBtn) bgPlayBtn.textContent = "播放背景音樂";
}

/**
 * 同時播放已載入的背景音樂，每軌套用延遲(秒)與音量 0–100%，循環播放
 * @param {number} [playbackOffsetSec] 當前 Master Timeline 播放位置（秒）
 */
function startBgMusic(playbackOffsetSec = 0) {
  let hasAny = false;
  for (let i = 0; i < 4; i += 1) {
    if (bgBuffers[i]) hasAny = true;
  }
  if (!hasAny) return;
  stopBgMusic();

  if (audioContext.state === "suspended") {
    audioContext.resume().catch(() => {});
  }

  if (!bgGainNode) {
    bgGainNode = audioContext.createGain();
    bgGainNode.gain.value = 0.35;
    bgGainNode.connect(audioContext.destination);
  }

  const now = audioContext.currentTime;
  // 若目前沒有任何軌在跑，BG 單獨播放時也要建立 Master Timeline
  if (!isPlaying && !isBgPlaying) {
    masterStartTime = now - playbackOffsetSec;
    contentStartTime = masterStartTime;
  }
  /** @type {(GainNode | null)[]} 依 slot 0..3 對應，方便即時音量對應 */
  bgTrackGainNodes = [null, null, null, null];
  for (let i = 0; i < 4; i += 1) {
    if (!bgBuffers[i]) continue;
    const delayEl = bgDelayInputs[i];
    const bgManualOffsetSec =
      delayEl && delayEl.value !== "" ? parseFloat(delayEl.value) || 0 : 0;
    const volEl = bgVolumeInputs[i];
    const volume = volEl ? (parseInt(volEl.value, 10) / 100) : 1;
    const buffer = bgBuffers[i];
    const bgAutoOffsetSec = bgAutoOffsetsSec[i] || 0;
    const bgFinalOffsetSec = bgAutoOffsetSec + bgManualOffsetSec;
    const trackGain = audioContext.createGain();
    trackGain.gain.value = volume;
    trackGain.connect(bgGainNode);
    bgTrackGainNodes[i] = trackGain;
    const source = audioContext.createBufferSource();
    source.buffer = buffer;
    source.loop = true;
    source.connect(trackGain);
    const timing = computeStartAtAndOutputTime(
      playbackOffsetSec,
      bgFinalOffsetSec,
      buffer.duration,
      true
    );
    source.start(now + timing.outputDelaySec, timing.startAtSec);
    bgSourceNodes.push(source);
  }
  isBgPlaying = true;
  if (bgPlayBtn) bgPlayBtn.textContent = "停止背景音樂";
}

/**
 * 更新背景音樂檔名顯示與狀態按鈕
 */
function updateBgMusicUI() {
  const count = bgBuffers.filter(Boolean).length;
  for (let i = 0; i < 4; i += 1) {
    const el = bgNameEls[i];
    if (el) el.textContent = bgFileNames[i] || "未選";
  }
  if (bgMusicStatusEl) {
    const names = bgFileNames.filter(Boolean);
    bgMusicStatusEl.textContent =
      count === 0 ? "未載入" : names.length ? `已載入：${names.join("、")}` : `已載入 ${count} 首`;
  }
  if (bgPlayBtn) {
    bgPlayBtn.disabled = !bgBuffers.some(Boolean);
  }
}

bgMusicInputs.forEach((inputEl, index) => {
  if (!inputEl) return;
  inputEl.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) {
      bgBuffers[index] = null;
      bgFileNames[index] = "";
      updateBgMusicUI();
      e.target.value = "";
      return;
    }
    try {
      const buf = await readFileAsArrayBuffer(file);
      const decoded = await decodeAudio(buf);
      bgBuffers[index] = decoded;
      bgFileNames[index] = file.name;
      updateBgMusicUI();
      if (bgAutoPlayCheckbox && bgAutoPlayCheckbox.checked) {
        const offset = isPlaying ? playbackOffsetSec : 0;
        startBgMusic(offset);
      }
    } catch (err) {
      console.error("背景音樂解碼失敗", err);
      if (bgMusicStatusEl) bgMusicStatusEl.textContent = "載入失敗，請重試";
    }
    e.target.value = "";
  });
});

if (bgPlayBtn) {
  bgPlayBtn.addEventListener("click", () => {
    if (audioContext.state === "suspended") {
      audioContext.resume().catch(() => {});
    }
    if (isBgPlaying) {
      stopBgMusic();
    } else {
      const offset = isPlaying ? playbackOffsetSec : 0;
      startBgMusic(offset);
    }
  });
}

if (refVolumeInput && refVolumeValEl) {
  refVolumeInput.addEventListener("input", () => {
    refVolumeValEl.textContent = refVolumeInput.value;
    if (refGainNode) refGainNode.gain.value = parseInt(refVolumeInput.value, 10) / 100;
  });
}
if (userVolumeInput && userVolumeValEl) {
  userVolumeInput.addEventListener("input", () => {
    userVolumeValEl.textContent = userVolumeInput.value;
    if (userGainNode) userGainNode.gain.value = parseInt(userVolumeInput.value, 10) / 100;
  });
}
bgVolumeInputs.forEach((inputEl, i) => {
  const valEl = bgVolumeValEls[i];
  if (inputEl && valEl) {
    inputEl.addEventListener("input", () => {
      valEl.textContent = inputEl.value;
      if (bgTrackGainNodes[i]) bgTrackGainNodes[i].gain.value = parseInt(inputEl.value, 10) / 100;
    });
  }
});

/**
 * RMS 自動對齊按鈕：以背景音樂（若有）或 ref 當 master，對 user 做 RMS alignment
 */
if (rmsAlignBtn) {
  rmsAlignBtn.addEventListener("click", async () => {
    if (!lastPlaybackData) {
      alert("請先載入並分析音檔。");
      return;
    }

    if (audioContext.state === "suspended") {
      try {
        await audioContext.resume();
      } catch {
        // ignore
      }
    }

    if (rmsAlignStatusEl) rmsAlignStatusEl.textContent = "RMS 對齊計算中...";

    const refBuffer = lastPlaybackData.refBuffer;
    const userBuffer = lastPlaybackData.userBuffer;

    /** Master Track 選擇規則：
     *  1) 優先使用鋼琴伴奏 masterPianoBuffer
     *  2) 否則，若存在任何 BG，使用第一條 BG
     *  3) 否則 fallback 為 REF
     */
    /** @type {AudioBuffer | null} */
    let masterBuffer = null;
    /** @type {"Piano"|"BG"|"REF"} */
    let masterKind = "REF";
    /** @type {number} BG master 的索引，僅在 masterKind === "BG" 時有意義 */
    let masterBgIndex = -1;

    if (masterPianoBuffer) {
      masterBuffer = masterPianoBuffer;
      masterKind = "Piano";
    } else {
      for (let i = 0; i < bgBuffers.length; i += 1) {
        if (bgBuffers[i]) {
          masterBuffer = bgBuffers[i];
          masterKind = "BG";
          masterBgIndex = i;
          break;
        }
      }
      if (!masterBuffer) {
        masterBuffer = refBuffer;
        masterKind = "REF";
      }
    }

    if (!masterBuffer || !refBuffer || !userBuffer) {
      if (rmsAlignStatusEl) rmsAlignStatusEl.textContent = "無法進行 RMS 對齊（缺少 master/ref/user 軌）。";
      return;
    }

    const hasBg = bgBuffers.some(Boolean);

    // -----------------------------
    // A) Onset-based auto-align (方案 A)
    // -----------------------------

    /**
     * band-pass + onset envelope + NCC 專用參數
     */
    const ONSET_ALIGN_HIGHPASS_HZ = 200;
    const ONSET_ALIGN_LOWPASS_HZ = 4000;
    const ONSET_ALIGN_FRAME_MS = 15;
    const ONSET_ALIGN_HOP_MS = 5;
    const MAX_ONSET_ALIGN_LAG_SECONDS = 2.0;
    const ONSET_ALIGN_CONF_THRESHOLD = 0.2;

    /**
     * voiced-aware refinement 入口微調參數
     */
    const MIN_VOICED_MS = 300;
    const REFINE_MAX_SEC = 0.12;
    const VOICE_REFINE_CONF_THRESHOLD = 0.2;
    const STABLE_LOGP_STD = 0.08;
    const VOICE_REFINE_MIN_COUNT_PER_LAG = 20;
    const VOICE_REFINE_MIN_COUNT_APPLY = 30;

    /**
     * 使用 OfflineAudioContext 做 band-pass（high-pass + low-pass），並輸出處理後 AudioBuffer
     * @param {AudioBuffer} source
     * @param {number} hpHz
     * @param {number} lpHz
     * @returns {Promise<AudioBuffer>}
     */
    async function renderBandpassOffline(source, hpHz, lpHz) {
      const ch = source.numberOfChannels;
      const len = source.length;
      const sr = source.sampleRate;
      const offline = new OfflineAudioContext(ch, len, sr);
      const src = offline.createBufferSource();
      src.buffer = source;
      const hp = offline.createBiquadFilter();
      hp.type = "highpass";
      hp.frequency.value = hpHz;
      const lp = offline.createBiquadFilter();
      lp.type = "lowpass";
      lp.frequency.value = lpHz;
      src.connect(hp);
      hp.connect(lp);
      lp.connect(offline.destination);
      src.start(0);
      return offline.startRendering();
    }

    /**
     * 由 band-passed buffer 計算 onset-like envelope：
     * - frame RMS
     * - 一階差分 + half-wave rectification
     * - 輕微平滑 + 有效段門檻
     * @param {AudioBuffer} buffer
     * @returns {{ env: number[]; hopSec: number }}
     */
    function computeOnsetEnvelopeForAlign(buffer) {
      const sr = buffer.sampleRate;
      const hopSamples = Math.max(1, Math.floor((ONSET_ALIGN_HOP_MS / 1000) * sr));
      const frameSamples = Math.max(1, Math.floor((ONSET_ALIGN_FRAME_MS / 1000) * sr));
      const chData = getMonoChannel(buffer);
      const n = chData.length;

      /** @type {number[]} */
      const rmsEnv = [];
      for (let start = 0; start + frameSamples <= n; start += hopSamples) {
        let sum = 0;
        for (let i = 0; i < frameSamples; i += 1) {
          const s = chData[start + i];
          sum += s * s;
        }
        const rms = Math.sqrt(sum / frameSamples);
        rmsEnv.push(rms);
      }

      if (!rmsEnv.length) {
        return { env: [], hopSec: ONSET_ALIGN_HOP_MS / 1000 };
      }

      /** onset-like: 一階差分 + half-wave rectification */
      /** @type {number[]} */
      const onset = [];
      for (let i = 0; i < rmsEnv.length; i += 1) {
        if (i === 0) {
          onset.push(0);
        } else {
          const d = rmsEnv[i] - rmsEnv[i - 1];
          onset.push(d > 0 ? d : 0);
        }
      }

      const smoothed = movingAverage(onset, 5);

      // 有效段門檻：低於一定比例視為 0，避免無聲段參與 correlation
      let maxVal = 0;
      for (let i = 0; i < smoothed.length; i += 1) {
        if (smoothed[i] > maxVal) maxVal = smoothed[i];
      }
      if (maxVal <= 0) {
        return { env: smoothed.map(() => 0), hopSec: ONSET_ALIGN_HOP_MS / 1000 };
      }
      const thr = maxVal * 0.1;
      for (let i = 0; i < smoothed.length; i += 1) {
        if (smoothed[i] < thr) smoothed[i] = 0;
      }

      return { env: smoothed, hopSec: ONSET_ALIGN_HOP_MS / 1000 };
    }

    /**
     * 方案 A：band-pass + onset envelope + NCC
     * @param {AudioBuffer} masterBuf
     * @param {AudioBuffer} targetBuf
     * @param {number} [maxShiftSec]
     * @returns {Promise<{ offsetSeconds: number; bestCorr: number }>}
     */
    async function onsetAutoAlign(masterBuf, targetBuf, maxShiftSec = MAX_ONSET_ALIGN_LAG_SECONDS) {
      const bpMaster = await renderBandpassOffline(
        masterBuf,
        ONSET_ALIGN_HIGHPASS_HZ,
        ONSET_ALIGN_LOWPASS_HZ
      );
      const bpTarget = await renderBandpassOffline(
        targetBuf,
        ONSET_ALIGN_HIGHPASS_HZ,
        ONSET_ALIGN_LOWPASS_HZ
      );

      const envA = computeOnsetEnvelopeForAlign(bpMaster);
      const envB = computeOnsetEnvelopeForAlign(bpTarget);

      let a = envA.env;
      let b = envB.env;
      const len = Math.min(a.length, b.length);
      if (!len) return { offsetSeconds: 0, bestCorr: 0 };
      a = a.slice(0, len);
      b = b.slice(0, len);

      const hopSec = envA.hopSec || (ONSET_ALIGN_HOP_MS / 1000);
      const maxLagFrames = Math.max(
        1,
        Math.min(Math.floor((maxShiftSec || MAX_ONSET_ALIGN_LAG_SECONDS) / hopSec), len - 1)
      );

      const aNorm = zNormalize(a);
      const bNorm = zNormalize(b);

      // 先用既有 crossCorrelateLimited 找整數 lag
      const coarse = crossCorrelateLimited(aNorm, bNorm, maxLagFrames);
      let bestLag = coarse.bestLag;
      let bestCorr = coarse.bestCorr;

      if (!Number.isFinite(bestCorr) || Number.isNaN(bestCorr)) {
        return { offsetSeconds: 0, bestCorr: 0 };
      }

      // 亞幀插值：用 parabolic interpolation 在 corr[k-1], corr[k], corr[k+1] 上微調 lag
      if (bestLag > -maxLagFrames && bestLag < maxLagFrames) {
        /**
         * 直接在這裡重算 k-1, k, k+1 的 NCC 值，避免修改 crossCorrelateLimited 實作。
         * 這三個值用來做二次曲線插值，得到子 frame 級別的 lag 修正。
         */
        function corrAtLag(lag) {
          let sumAB = 0;
          let sumA2 = 0;
          let sumB2 = 0;
          let count = 0;

          if (lag >= 0) {
            for (let i = 0; i < len - lag; i += 1) {
              const va = aNorm[i];
              const vb = bNorm[i + lag];
              sumAB += va * vb;
              sumA2 += va * va;
              sumB2 += vb * vb;
              count += 1;
            }
          } else {
            const k = -lag;
            for (let i = 0; i < len - k; i += 1) {
              const va = aNorm[i + k];
              const vb = bNorm[i];
              sumAB += va * vb;
              sumA2 += va * va;
              sumB2 += vb * vb;
              count += 1;
            }
          }

          if (count === 0) return 0;
          const denom = Math.sqrt(sumA2 * sumB2) || 1;
          return sumAB / denom;
        }

        const k = bestLag;
        const c1 = corrAtLag(k - 1);
        const c2 = bestCorr;
        const c3 = corrAtLag(k + 1);

        const denom = (c1 - 2 * c2 + c3) || 0;
        if (denom !== 0) {
          const delta =
            0.5 * (c1 - c3) / denom; // 典型 parabolic interpolation 修正量，單位：frame
          const refinedLag = k + delta;
          const refinedOffset = refinedLag * hopSec;
          // confidence 仍用整數 lag 的 bestCorr 即可
          return { offsetSeconds: refinedOffset, bestCorr };
        }
      }

      const offsetSeconds = bestLag * hopSec;
      return { offsetSeconds, bestCorr };
    }

    /**
     * 在 pitch series 中找第一段足夠長且穩定的 voiced 片段
     * @param {number[]} pitches
     * @param {number} hopSec
     * @returns {{ start: number; end: number } | null}
     */
    function findFirstVoicedSegment(pitches, hopSec) {
      const minFrames = Math.max(1, Math.floor((MIN_VOICED_MS / 1000) / hopSec));
      const n = pitches.length;

      let runStart = -1;
      for (let i = 0; i <= n; i += 1) {
        const p = i < n ? pitches[i] : 0;
        const voiced = p && Number.isFinite(p) && p > 0;
        if (voiced) {
          if (runStart === -1) runStart = i;
        } else if (runStart !== -1) {
          const runEnd = i;
          const len = runEnd - runStart;
          if (len >= minFrames) {
            // 檢查此片段的 log2(pitch) 標準差是否足夠小（避免滑音/抖動過大）
            let sum = 0;
            let sumSq = 0;
            let count = 0;
            for (let k = runStart; k < runEnd; k += 1) {
              const pk = pitches[k];
              if (!pk || !Number.isFinite(pk) || pk <= 0) continue;
              const lp = Math.log2(pk);
              if (!Number.isFinite(lp)) continue;
              sum += lp;
              sumSq += lp * lp;
              count += 1;
            }
            if (count > 1) {
              const mean = sum / count;
              const varVal = Math.max(0, sumSq / count - mean * mean);
              const std = Math.sqrt(varVal);
              if (std <= STABLE_LOGP_STD) {
                return { start: runStart, end: runEnd };
              }
            }
          }
          runStart = -1;
        }
      }
      return null;
    }

    /**
     * 入口人聲微調：只在 user 第一段 voiced 片段附近做小範圍 NCC
     * coarseOffsetSec > 0 代表 USER 晚於 REF
     * @param {number[]} refPitches
     * @param {number[]} userPitches
     * @param {number} coarseOffsetSec user 相對 ref 的粗略 offset
     * @returns {{ refineDeltaSec: number; bestCorr: number }}
     */
    function voicedRefineAlign(refPitches, userPitches, coarseOffsetSec) {
      const hopSec = SAMPLE_INTERVAL_SECONDS;
      const n = Math.min(refPitches.length, userPitches.length);
      if (!n) return { refineDeltaSec: 0, bestCorr: 0 };

      const voicedSeg = findFirstVoicedSegment(userPitches.slice(0, n), hopSec);
      if (!voicedSeg) return { refineDeltaSec: 0, bestCorr: 0 };

      const u0 = voicedSeg.start;
      const u1 = voicedSeg.end;
      const winLen = u1 - u0;
      if (winLen < 2) return { refineDeltaSec: 0, bestCorr: 0 };

      const maxLagFrames = Math.max(1, Math.floor(REFINE_MAX_SEC / hopSec));
      const coarseLagFrames = Math.round(coarseOffsetSec / hopSec);

      let bestLagFrames = coarseLagFrames;
      let bestCorr = -Infinity;
      let bestCount = 0;

      // 以 coarseLag 為中心，在 ±maxLagFrames 範圍內掃描
      for (let d = -maxLagFrames; d <= maxLagFrames; d += 1) {
        const lagFrames = coarseLagFrames + d;

        let sumA = 0;
        let sumB = 0;
        let sumA2 = 0;
        let sumB2 = 0;
        let sumAB = 0;
        let count = 0;

        for (let j = 0; j < winLen; j += 1) {
          const ui = u0 + j;
          const ri = ui - lagFrames;
          if (ri < 0 || ri >= n) continue;

          const pRef = refPitches[ri];
          const pUser = userPitches[ui];
          if (!pRef || !pUser || !Number.isFinite(pRef) || !Number.isFinite(pUser)) continue;

          const x = Math.log2(pRef);
          const y = Math.log2(pUser);
          if (!Number.isFinite(x) || !Number.isFinite(y)) continue;

          sumA += x;
          sumB += y;
          sumA2 += x * x;
          sumB2 += y * y;
          sumAB += x * y;
          count += 1;
        }

        if (count < VOICE_REFINE_MIN_COUNT_PER_LAG) continue;

        const meanA = sumA / count;
        const meanB = sumB / count;
        const num = sumAB - count * meanA * meanB;
        const varA = sumA2 - count * meanA * meanA;
        const varB = sumB2 - count * meanB * meanB;
        const den = Math.sqrt(varA * varB);
        const corr = den > 0 ? num / den : 0;
        if (corr > bestCorr) {
          bestCorr = corr;
          bestLagFrames = lagFrames;
          bestCount = count;
        }
      }

      if (!Number.isFinite(bestCorr) || bestCorr <= 0 || bestCount < VOICE_REFINE_MIN_COUNT_APPLY) {
        return { refineDeltaSec: 0, bestCorr: 0 };
      }

      const refineFrames = bestLagFrames - coarseLagFrames;
      let refineDeltaSec = refineFrames * hopSec;
      if (refineDeltaSec > REFINE_MAX_SEC) refineDeltaSec = REFINE_MAX_SEC;
      else if (refineDeltaSec < -REFINE_MAX_SEC) refineDeltaSec = -REFINE_MAX_SEC;

      return { refineDeltaSec, bestCorr };
    }

    // 小工具：給定自動對齊結果與對應 buffer，計算 Attack 微調與最終 offset
    /**
     * @param {{ offsetSeconds: number; bestCorr: number }} auto
     * @param {AudioBuffer} masterBuf
     * @param {AudioBuffer} trackBuf
     */
    function computeFinalOffsetWithAttack(auto, masterBuf, trackBuf) {
      const rmsOffset = auto.offsetSeconds;
      const rmsCorr = auto.bestCorr;
      let apply = false;
      let attackOffset = 0;
      let attackConf = 0;
      let finalOffset = rmsOffset;

      // RMS 信心不足：只回傳參考資訊，不套用
      if (rmsCorr < 0.15) {
        return { rmsOffset, rmsCorr, attackOffset, attackConf, finalOffset, apply };
      }

      // RMS OK，先採用 RMS offset
      apply = true;
      finalOffset = rmsOffset;

      // C) Attack Alignment 只做微調，且有信心門檻與防過修保護
      const attack = attackAutoAlign(masterBuf, trackBuf);
      const MAX_ATTACK_ADJUST = 0.12;
      // clamp attack 微調量，避免推太多
      const clampedAttack = Math.max(
        -MAX_ATTACK_ADJUST,
        Math.min(MAX_ATTACK_ADJUST, attack.attackOffsetSec)
      );
      attackOffset = clampedAttack;
      attackConf = attack.userAttack.conf;

      const ATTACK_CONF_THRESHOLD = 4.0;
      if (attackConf >= ATTACK_CONF_THRESHOLD &&
          Math.abs(clampedAttack) <= Math.abs(rmsOffset) * 0.5) {
        finalOffset = rmsOffset + clampedAttack;
      } else {
        // conf 不足，不套用 Attack，只維持 RMS
        attackOffset = 0;
      }

      return { rmsOffset, rmsCorr, attackOffset, attackConf, finalOffset, apply };
    }

    const statusParts = [];
    statusParts.push(`Master: ${masterKind}`);

    // 2️⃣ 若 master 存在（Piano 或 BG 或 REF）：分別計算 REF/USER 對 master 的 onset-based auto align
    const refAuto = await onsetAutoAlign(masterBuffer, refBuffer, MAX_ONSET_ALIGN_LAG_SECONDS);
    const userAuto = await onsetAutoAlign(masterBuffer, userBuffer, MAX_ONSET_ALIGN_LAG_SECONDS);

    // 3️⃣ 入口人聲微調：以 pitch series 為基礎，只調整 USER 相對 REF 的 offset
    let finalRefAutoOffsetSec = refAuto.offsetSeconds;
    let finalUserAutoOffsetSec = userAuto.offsetSeconds;
    let voiceRefineInfo = null;

    if (lastPlaybackData && lastPlaybackData.refPitches && lastPlaybackData.userPitches) {
      const coarseUserVsRefSec = userAuto.offsetSeconds - refAuto.offsetSeconds;
      const refine = voicedRefineAlign(
        lastPlaybackData.refPitches,
        lastPlaybackData.userPitches,
        coarseUserVsRefSec
      );
      voiceRefineInfo = refine;

      if (
        Number.isFinite(refine.refineDeltaSec) &&
        Math.abs(refine.refineDeltaSec) <= REFINE_MAX_SEC &&
        refine.bestCorr >= VOICE_REFINE_CONF_THRESHOLD
      ) {
        const refinedUserVsRefSec = coarseUserVsRefSec + refine.refineDeltaSec;
        finalUserAutoOffsetSec = finalRefAutoOffsetSec + refinedUserVsRefSec;
      }
    }

    const refAutoForAttack = { offsetSeconds: finalRefAutoOffsetSec, bestCorr: refAuto.bestCorr };
    const userAutoForAttack = { offsetSeconds: finalUserAutoOffsetSec, bestCorr: userAuto.bestCorr };

    const refInfo = computeFinalOffsetWithAttack(refAutoForAttack, masterBuffer, refBuffer);
    const userInfo = computeFinalOffsetWithAttack(userAutoForAttack, masterBuffer, userBuffer);

    if (lastPlaybackData) {
      if (refInfo.apply) lastPlaybackData.refOffsetSec = refInfo.finalOffset;
      if (userInfo.apply) lastPlaybackData.userOffsetSec = userInfo.finalOffset;
    }

    const refOffRms = refInfo.rmsOffset.toFixed(3);
    const refCorr = refInfo.rmsCorr.toFixed(3);
    const refAtk = refInfo.attackOffset.toFixed(3);
    const refConf = refInfo.attackConf.toFixed(2);
    const refFinal = refInfo.finalOffset.toFixed(3);
    statusParts.push(
      `REF | RMS: ${refOffRms}s (corr ${refCorr}), Attack: ${refAtk}s (conf ${refConf}), Final: ${refFinal}s`
    );

    const userOffRms = userInfo.rmsOffset.toFixed(3);
    const userCorr = userInfo.rmsCorr.toFixed(3);
    const userAtk = userInfo.attackOffset.toFixed(3);
    const userConf = userInfo.attackConf.toFixed(2);
    const userFinal = userInfo.finalOffset.toFixed(3);
    if (voiceRefineInfo) {
      const refineDelta = voiceRefineInfo.refineDeltaSec.toFixed(3);
      const refineConf = voiceRefineInfo.bestCorr.toFixed(3);
      statusParts.push(
        `USER | RMS: ${userOffRms}s (corr ${userCorr}), VoiceRefine: ${refineDelta}s (conf ${refineConf}), Attack: ${userAtk}s (conf ${userConf}), Final: ${userFinal}s`
      );
    } else {
      statusParts.push(
        `USER | RMS: ${userOffRms}s (corr ${userCorr}), Attack: ${userAtk}s (conf ${userConf}), Final: ${userFinal}s`
      );
    }

    // BG：若存在，逐條計算對 master 的 onset-based auto offset，僅 auto，不做 Attack
    const bgPieces = [];
    if (hasBg) {
      bgAutoOffsetsSec = [0, 0, 0, 0];
      for (let i = 0; i < bgBuffers.length; i += 1) {
        const buf = bgBuffers[i];
        if (!buf) continue;
        // 若此 BG 即為 master BG，本身 offset=0
        if (masterKind === "BG" && i === masterBgIndex) {
          bgAutoOffsetsSec[i] = 0;
          bgPieces.push(`BG${i + 1} | AUTO: 0.000s (conf 1.00), Final: 0.000s`);
          continue;
        }
        const bgAuto = await onsetAutoAlign(masterBuffer, buf, MAX_ONSET_ALIGN_LAG_SECONDS);
        if (bgAuto.bestCorr >= ONSET_ALIGN_CONF_THRESHOLD) {
          bgAutoOffsetsSec[i] = bgAuto.offsetSeconds;
        }
        const bgOff = bgAuto.offsetSeconds.toFixed(3);
        const bgCorr = bgAuto.bestCorr.toFixed(3);
        const bgFinal = bgAutoOffsetsSec[i].toFixed(3);
        bgPieces.push(`BG${i + 1} | AUTO: ${bgOff}s (conf ${bgCorr}), Final: ${bgFinal}s`);
      }
      if (bgPieces.length) {
        statusParts.push(bgPieces.join(" ; "));
      }
    }

    const lowConfidence =
      refInfo.rmsCorr < ONSET_ALIGN_CONF_THRESHOLD || userInfo.rmsCorr < ONSET_ALIGN_CONF_THRESHOLD;

    if (rmsAlignStatusEl) {
      rmsAlignStatusEl.textContent =
        statusParts.join("  ||  ") +
          (lowConfidence ? "（自動對齊信心低，建議手動微調）" : "");
    }

    // 若目前正在播放，立即套用新的 offset（重新 start 同步播放）
    if (isPlaying) {
      stopPlayback();
      startSyncedPlayback();
    }
  });
}

/**
 * 根據 pitch 序列中的無聲區段自動偵測分句
 * @param {number[]} pitches
 * @param {number} durationSec
 * @returns {{ id: number; startSec: number; endSec: number }[]}
 */
function detectPhrases(pitches, durationSec) {
  /** @type {{ id: number; startSec: number; endSec: number }[]} */
  const result = [];
  const hop = SAMPLE_INTERVAL_SECONDS;
  const gapFrames = Math.floor((PHRASE_GAP_MS / 1000) / hop);
  const minFrames = Math.floor((PHRASE_MIN_MS / 1000) / hop);

  if (!pitches.length || durationSec <= 0) {
    return result;
  }

  let startIdx = 0;
  let silentCount = 0;

  for (let i = 0; i < pitches.length; i += 1) {
    const p = pitches[i];
    const valid = p && Number.isFinite(p) && p > 0;

    if (!valid) {
      silentCount += 1;
    } else {
      if (silentCount >= gapFrames && gapFrames > 0) {
        const endIdx = i - silentCount;
        if (endIdx - startIdx >= minFrames) {
          result.push({
            id: result.length + 1,
            startSec: startIdx * hop,
            endSec: endIdx * hop,
          });
        }
        startIdx = i;
      }
      silentCount = 0;
    }
  }

  // 收尾：最後一段句子
  if (pitches.length - startIdx >= minFrames) {
    result.push({
      id: result.length + 1,
      startSec: startIdx * hop,
      endSec: durationSec,
    });
  }

  return result;
}

/**
 * 渲染分句 Phrase Bar：依照 phrases 與當前曲目長度產生可點擊區塊
 */
function renderPhraseBar() {
  if (!phraseBarEl) return;
  if (!lastPlaybackData) {
    phraseBarEl.innerHTML = "";
    return;
  }

  const duration = lastPlaybackData.durationSec || 0;
  if (!duration || duration <= 0) {
    phraseBarEl.innerHTML = "";
    return;
  }

  phraseBarEl.innerHTML = "";

  phrases.forEach((p) => {
    const seg = document.createElement("div");
    const widthPct = Math.max(
      0.5,
      ((p.endSec - p.startSec) / duration) * 100
    );
    seg.style.flex = `0 0 ${widthPct}%`;
    seg.className = "phrase-seg";
    seg.dataset.start = String(p.startSec);
    seg.dataset.id = String(p.id);

    seg.addEventListener("click", () => {
      if (!lastPlaybackData) return;
      const dur = lastPlaybackData.durationSec || 0;
      const wasPlaying = isPlaying;
      if (wasPlaying) stopPlayback();

      const clampedStart = Math.max(0, Math.min(p.startSec, dur));
      playbackOffsetSec = clampedStart;
      if (dur > 0) {
        seekToPercent(clampedStart / dur);
      }

      if (wasPlaying) {
        startSyncedPlayback();
      }
    });

    // 左右拖曳 handle
    const leftHandle = document.createElement("div");
    leftHandle.className = "phrase-handle left";
    leftHandle.addEventListener("mousedown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      if (isProRecording) return;
      if (!lastPlaybackData || !phraseBarEl) return;
      phraseDragIndex = phrases.indexOf(p);
      phraseDragIsLeft = true;
      phraseDragStartX = e.clientX;
      phraseDragStartStartSec = p.startSec;
      phraseDragStartEndSec = p.endSec;
      const rect = phraseBarEl.getBoundingClientRect();
      phraseDragTimelineWidth = rect.width || 1;
    });

    const rightHandle = document.createElement("div");
    rightHandle.className = "phrase-handle right";
    rightHandle.addEventListener("mousedown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      if (isProRecording) return;
      if (!lastPlaybackData || !phraseBarEl) return;
      phraseDragIndex = phrases.indexOf(p);
      phraseDragIsLeft = false;
      phraseDragStartX = e.clientX;
      phraseDragStartStartSec = p.startSec;
      phraseDragStartEndSec = p.endSec;
      const rect = phraseBarEl.getBoundingClientRect();
      phraseDragTimelineWidth = rect.width || 1;
    });

    seg.appendChild(leftHandle);
    seg.appendChild(rightHandle);

    phraseBarEl.appendChild(seg);
  });
}

/**
 * 清理錄音相關節點
 */
function cleanupRecordingNodes() {
  if (scriptNode) {
    try {
      scriptNode.disconnect();
    } catch {
      // ignore
    }
    scriptNode.onaudioprocess = null;
    scriptNode = null;
  }

  if (mediaStream) {
    mediaStream.getTracks().forEach((t) => t.stop());
    mediaStream = null;
  }
}

/**
 * 停止錄音並建立 AudioBuffer
 */
function stopRecording() {
  if (!isRecording) return;
  isRecording = false;

  cleanupRecordingNodes();

  // 合併 chunks
  const totalLength = recordedChunks.reduce(
    (sum, ch) => sum + ch.length,
    0
  );
  const merged = new Float32Array(totalLength);
  let offset = 0;
  for (const ch of recordedChunks) {
    merged.set(ch, offset);
    offset += ch.length;
  }
  recordedChunks = [];

  if (totalLength > 0) {
    const buffer = audioContext.createBuffer(
      1,
      totalLength,
      audioContext.sampleRate
    );
    buffer.copyToChannel(merged, 0, 0);
    recordedUserAudio = buffer;
    // 錄音完成視為更換 user 音檔來源，同樣需要重置分析/播放狀態
    invalidateAnalysisState();
    if (recordStatusEl) {
      recordStatusEl.textContent =
        "錄音完成：將以這段麥克風錄音當作你的演唱。";
    }
  } else if (recordStatusEl) {
    recordStatusEl.textContent = "錄音太短或沒有聲音，請再試一次。";
  }

  if (recordBtn) {
    recordBtn.disabled = false;
    recordBtn.textContent = `用麥克風錄一段（${RECORD_SECONDS} 秒）`;
  }
}

/**
 * 開始以麥克風錄音固定秒數
 */
async function startRecording() {
  if (isRecording) return;

  try {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      alert("此瀏覽器不支援麥克風錄音。");
      return;
    }

    mediaStream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const source = audioContext.createMediaStreamSource(mediaStream);

    const bufferSize = 2048;
    scriptNode = audioContext.createScriptProcessor(bufferSize, 1, 1);
    recordedChunks = [];
    const startTime = audioContext.currentTime;
    isRecording = true;

    if (recordStatusEl) {
      recordStatusEl.textContent = "錄音中... 請開始演唱。";
    }
    if (recordBtn) {
      recordBtn.disabled = true;
      recordBtn.textContent = "錄音中（請稍候）...";
    }

    scriptNode.onaudioprocess = (event) => {
      if (!isRecording) return;
      const input = event.inputBuffer.getChannelData(0);
      recordedChunks.push(new Float32Array(input));

      const elapsed = audioContext.currentTime - startTime;
      if (elapsed >= RECORD_SECONDS) {
        stopRecording();
      }
    };

    source.connect(scriptNode);
    scriptNode.connect(audioContext.destination);
  } catch (err) {
    console.error("啟動錄音失敗", err);
    alert("啟動錄音失敗，請確認已允許麥克風權限。");
    cleanupRecordingNodes();
  }
}

if (recordBtn) {
  recordBtn.addEventListener("click", () => {
    if (audioContext.state === "suspended") {
      audioContext.resume().catch(() => {});
    }
    startRecording();
  });
}

// 單一「開始練習」按鈕：第一次點擊開始錄音，再次點擊停止並顯示預覽
if (startRecordBtn) {
  startRecordBtn.addEventListener("click", async () => {
    const cap = getRecordingCapability();
    if (!cap.isSecureContext) {
      setDeviceWarning(
        "error",
        "需要在 https 或 http://localhost 才能錄音（請勿用 file:// 開啟）"
      );
      return;
    }
    if (!cap.hasGetUserMedia) {
      setDeviceWarning(
        "error",
        "此環境不支援麥克風 API（getUserMedia）。請改用 Chrome/Edge/Safari 正式瀏覽器。"
      );
      return;
    }
    if (cap.hasMediaRecorder) {
      setDeviceWarning(
        "warning",
        "錄音引擎：MediaRecorder（opus/webm 優先，可長錄音）"
      );
    } else {
      setDeviceWarning(
        "warning",
        "錄音引擎：MediaRecorder 不可用，將改用 WAV Fallback（檔案較大）。"
      );
    }

    if (audioContext.state === "suspended") {
      await audioContext.resume().catch(() => {});
    }
    if (!isProRecording && recUiState === "idle") {
      await startProRecording();
    }
  });
}

if (stopRecordBtn) {
  stopRecordBtn.addEventListener("click", async () => {
    if (isProRecording || recEngine !== "none") {
      await stopProRecording();
    }
  });
}

// 初始化錄音模式提示 / 裝置偵測
checkAudioDevices().catch(() => {});

if (recordModeEl) {
  recordModeEl.addEventListener("change", () => {
    checkAudioDevices().catch(() => {});
  });
}

if (previewPlayBtn) {
  previewPlayBtn.addEventListener("click", () => {
    if (pendingRecording) {
      playPreview(pendingRecording.buffer);
    }
  });
}

if (useAndCompareBtn) {
  useAndCompareBtn.addEventListener("click", async () => {
    await usePendingRecordingAndCompare();
  });
}

if (saveAsBtn) {
  saveAsBtn.addEventListener("click", () => {
    if (!pendingRecording) return;
    const base = pendingRecording;
    const dur =
      (base.durationSec || base.buffer?.duration) || 0;
    const rec = {
      id: nextRecordingId++,
      buffer: base.buffer,
      pitches: base.pitches ? base.pitches.slice() : [],
      createdAt: base.createdAt,
      isLongTake: dur > MAX_COMPARE_SEC,
    };
    recordings.push(rec);
    if (takeSelectorEl) {
      const option = document.createElement("option");
      option.value = String(rec.id);
      option.textContent = `Take ${rec.id}${rec.isLongTake ? " (未比對)" : ""}`;
      takeSelectorEl.appendChild(option);
      takeSelectorEl.value = String(rec.id);
    }
    alert("已另存新檔。若超過 30 秒，請改用上傳音檔方式進行比對。");
    clearPendingRecording();
  });
}

if (deletePendingBtn) {
  deletePendingBtn.addEventListener("click", () => {
    clearPendingRecording();
  });
}

if (startNewTakeBtn) {
  startNewTakeBtn.addEventListener("click", async () => {
    clearPendingRecording();
    await startProRecording();
  });
}

if (clearSettingsBtn) {
  clearSettingsBtn.addEventListener("click", () => {
    resetAllSessionState();
  });
}

if (replayBtn) {
  replayBtn.addEventListener("click", () => {
    replayWithCurrentSettings();
  });
  updateReplayButtonState();
}

if (takeSelectorEl) {
  takeSelectorEl.addEventListener("change", () => {
    if (!lastPlaybackData) return;
    const id = Number(takeSelectorEl.value);
    const rec = recordings.find((r) => r.id === id);
    if (!rec) return;

    if (!rec.pitches.length) {
      rec.pitches = analyzePitchFromBuffer(rec.buffer);
    }

    const aligned = alignPitchSeries(lastPlaybackData.refPitches, rec.pitches);

    lastPlaybackData = {
      ...lastPlaybackData,
      userBuffer: rec.buffer,
      userPitches: aligned.user,
      userOffsetSec: aligned.userOffsetSec ?? 0,
    };

    drawPitchCurves(
      lastPlaybackData.refPitches,
      lastPlaybackData.userPitches,
      undefined,
      lastPlaybackData.durationSec
    );
  });
}

/**
 * 上下分割拖曳：調整 top 高度（%）
 */
const resizerEl = document.getElementById("resizer");
if (resizerEl) {
  let resizing = false;
  resizerEl.addEventListener("mousedown", (e) => {
    e.preventDefault();
    resizing = true;
  });
  document.addEventListener("mousemove", (e) => {
    if (!resizing) return;
    const pct = (e.clientY / window.innerHeight) * 100;
    const clamped = Math.max(15, Math.min(85, pct));
    document.documentElement.style.setProperty("--top-height", String(clamped));
  });
  document.addEventListener("mouseup", () => {
    resizing = false;
  });
}