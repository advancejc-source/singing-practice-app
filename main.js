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
/** @type {HTMLButtonElement | null} */
const recordBtn = document.getElementById("record-btn");
/** @type {HTMLDivElement | null} */
const recordStatusEl = document.getElementById("record-status");
/** @type {HTMLDivElement | null} */
const segmentScoresEl = document.getElementById("segment-scores");
/** @type {HTMLDivElement | null} */
const pitchTooltipEl = document.getElementById("pitch-tooltip");

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
 * @type {{ refBuffer: AudioBuffer; userBuffer: AudioBuffer; refPitches: number[]; userPitches: number[]; durationSec: number } | null}
 */
let lastPlaybackData = null;

/** @type {AudioBufferSourceNode | null} */
let refSourceNode = null;
/** @type {AudioBufferSourceNode | null} */
let userSourceNode = null;
/** @type {number} */
let playStartTime = 0;
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
 * 內容時間軸的牆鐘起點（用於延遲時正確計算 elapsed）
 * @type {number}
 */
let contentStartTime = 0;

/** @type {MediaStream | null} */
let mediaStream = null;
/** @type {ScriptProcessorNode | null} */
let scriptNode = null;
/** @type {Float32Array[]} */
let recordedChunks = [];
/** @type {boolean} */
let isRecording = false;

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

/** 音準偏差圖：顏色區間（|deviationCents|）≤20 綠、20–50 黃、>50 紅 */
const DEVIATION_GREEN = 20;
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
 * 依偏差絕對值回傳 stroke 顏色（綠/黃/紅）
 * @param {number} absCents
 * @returns {string}
 */
function deviationColor(absCents) {
  if (absCents <= DEVIATION_GREEN) return "#22c55e";
  if (absCents <= DEVIATION_YELLOW) return "#eab308";
  return "#ef4444";
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

  const len = Math.min(refPitches.length, userPitches.length);
  if (!len) {
    lastPitchDrawData = null;
    return;
  }

  const refS = smoothPitches(refPitches);
  const userS = smoothPitches(userPitches);

  /** @type {(number | null)[]} 無聲/無效處為 null，斷線不畫到底 */
  const rawDeviations = [];
  for (let i = 0; i < len; i += 1) {
    const fr = refS[i];
    const fu = userS[i];
    if (!fr || fr <= 0 || !fu || fu <= 0) {
      rawDeviations.push(null);
    } else {
      const dev = 1200 * Math.log2(fu / fr);
      rawDeviations.push(Number.isFinite(dev) ? dev : null);
    }
  }
  const deviations = smoothDeviations(rawDeviations);

  const paddingLeft = 36;
  const paddingRight = 8;
  const paddingTop = 12;
  const paddingBottom = 24;
  const w = cssWidth - paddingLeft - paddingRight;
  const h = cssHeight - paddingTop - paddingBottom;

  /** Y 軸範圍：±cents，0 在中央 */
  const maxCents = 80;
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

  lastPitchDrawData = { deviations, len, paddingLeft, w, durationSec: totalSec };

  // 背景
  ctx.fillStyle = "#020617";
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  // 水平網格：±50c, ±20c, 0
  ctx.strokeStyle = "#1e293b";
  ctx.lineWidth = 1;
  const gridCents = [50, 20, 0, -20, -50];
  for (const c of gridCents) {
    const y = yForCents(c);
    ctx.beginPath();
    ctx.moveTo(paddingLeft, y);
    ctx.lineTo(cssWidth - paddingRight, y);
    ctx.stroke();
  }

  // Y 軸刻度：+50c / +20c / 0 / -20c / -50c（不顯示音名）
  ctx.fillStyle = "#94a3b8";
  ctx.font = "12px system-ui";
  ctx.textAlign = "right";
  ctx.textBaseline = "middle";
  for (const c of gridCents) {
    const y = yForCents(c);
    const label = c === 0 ? "0" : (c > 0 ? "+" : "") + c + "c";
    ctx.fillText(label, paddingLeft - 6, y);
  }

  // 基準線 y=0（標準 / In Tune）：加粗 + 輕微 glow
  const zeroY = yForCents(0);
  ctx.shadowColor = "rgba(56, 189, 248, 0.6)";
  ctx.shadowBlur = 8;
  ctx.strokeStyle = "#38bdf8";
  ctx.lineWidth = 2.5;
  ctx.beginPath();
  ctx.moveTo(paddingLeft, zeroY);
  ctx.lineTo(cssWidth - paddingRight, zeroY);
  ctx.stroke();
  ctx.shadowBlur = 0;
  ctx.shadowColor = "transparent";

  // 使用者偏差曲線：依 |deviation| 變色，null 斷線
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

    // 7. 合併功能：比對完自動開始播放並顯示動畫（含背景音樂，由 startSyncedPlayback 一併啟動）
    startSyncedPlayback();
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
 * 停止同步播放並還原 UI
 */
function stopPlayback() {
  if (!isPlaying) return;
  isPlaying = false;

  if (refSourceNode) {
    try {
      refSourceNode.stop();
      refSourceNode.disconnect();
    } catch (_) {}
    refSourceNode = null;
  }
  if (userSourceNode) {
    try {
      userSourceNode.stop();
      userSourceNode.disconnect();
    } catch (_) {}
    userSourceNode = null;
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
  const elapsed = Math.max(0, audioContext.currentTime - contentStartTime);
  const currentSec = playbackOffsetSec + elapsed;
  const progress = Math.min(1, Math.max(0, currentSec / lastPlaybackData.durationSec));

  drawPitchCurves(
    lastPlaybackData.refPitches,
    lastPlaybackData.userPitches,
    progress,
    lastPlaybackData.durationSec
  );

  updatingTimelineFromPlayback = true;
  if (timelineEl) timelineEl.value = String(progress * 100);
  if (timeCurrentEl) timeCurrentEl.textContent = formatTime(Math.min(currentSec, lastPlaybackData.durationSec));
  updatingTimelineFromPlayback = false;

  if (progress >= 1) {
    playbackOffsetSec = lastPlaybackData.durationSec;
    stopPlayback();
    return;
  }
  playbackAnimationId = requestAnimationFrame(animatePlayhead);
}

/**
 * 標準音與錄音檔同時開始播放，套用延遲(秒)、音量與時間軸起點（seek），並同步更新音高圖與時間條
 */
function startSyncedPlayback() {
  if (!lastPlaybackData || isPlaying) return;

  if (audioContext.state === "suspended") {
    audioContext.resume().catch(() => {});
  }

  const refDelaySec = refDelayInput && refDelayInput.value !== "" ? parseFloat(refDelayInput.value) || 0 : 0;
  const userDelaySec = userDelayInput && userDelayInput.value !== "" ? parseFloat(userDelayInput.value) || 0 : 0;
  const refVol = refVolumeInput ? (parseInt(refVolumeInput.value, 10) / 100) : 1;
  const userVol = userVolumeInput ? (parseInt(userVolumeInput.value, 10) / 100) : 1;

  refGainNode = audioContext.createGain();
  refGainNode.gain.value = refVol;
  refGainNode.connect(audioContext.destination);
  userGainNode = audioContext.createGain();
  userGainNode.gain.value = userVol;
  userGainNode.connect(audioContext.destination);

  refSourceNode = audioContext.createBufferSource();
  refSourceNode.buffer = lastPlaybackData.refBuffer;
  refSourceNode.connect(refGainNode);

  userSourceNode = audioContext.createBufferSource();
  userSourceNode.buffer = lastPlaybackData.userBuffer;
  userSourceNode.connect(userGainNode);

  playStartTime = audioContext.currentTime;
  const refBufOffset = (lastPlaybackData.refOffsetSec ?? 0) + playbackOffsetSec;
  const userBufOffset = (lastPlaybackData.userOffsetSec ?? 0) + playbackOffsetSec;
  refSourceNode.start(playStartTime + refDelaySec, refBufOffset);
  userSourceNode.start(playStartTime + userDelaySec, userBufOffset);
  contentStartTime = playStartTime + Math.min(refDelaySec, userDelaySec);

  isPlaying = true;
  if (analyzeBtn) analyzeBtn.textContent = "停止播放";
  if (bgBuffers.some(Boolean)) startBgMusic(playbackOffsetSec);
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
  if (isPlaying) {
    stopPlayback();
    startSyncedPlayback();
  } else {
    drawPitchCurves(
      lastPlaybackData.refPitches,
      lastPlaybackData.userPitches,
      pct,
      lastPlaybackData.durationSec
    );
  }
}

if (timelineEl) {
  timelineEl.addEventListener("input", () => {
    if (updatingTimelineFromPlayback) return;
    const pct = parseFloat(timelineEl.value) / 100;
    seekToPercent(pct);
  });
}

/** Pitch 偏差圖 tooltip：時間、偏差(± cents)、偏高/偏低，不顯示 Hz */
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
    const { paddingLeft, w, len, deviations, durationSec } = lastPitchDrawData;
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
      const dir = dev > 0 ? "偏高" : dev < 0 ? "偏低" : "唱準";
      pitchTooltipEl.innerHTML = `${timeStr}<br>${sign}${Math.round(dev)} cents · ${dir}`;
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
 * @param {number} [offsetSec] 從第幾秒開始播（與主播放時間軸同步時傳入 playbackOffsetSec）
 */
function startBgMusic(offsetSec = 0) {
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

  const baseTime = audioContext.currentTime;
  /** @type {(GainNode | null)[]} 依 slot 0..3 對應，方便即時音量對應 */
  bgTrackGainNodes = [null, null, null, null];
  for (let i = 0; i < 4; i += 1) {
    if (!bgBuffers[i]) continue;
    const delayEl = bgDelayInputs[i];
    const delaySec = delayEl && delayEl.value !== "" ? parseFloat(delayEl.value) || 0 : 0;
    const volEl = bgVolumeInputs[i];
    const volume = volEl ? (parseInt(volEl.value, 10) / 100) : 1;
    const buffer = bgBuffers[i];
    const trackGain = audioContext.createGain();
    trackGain.gain.value = volume;
    trackGain.connect(bgGainNode);
    bgTrackGainNodes[i] = trackGain;
    const source = audioContext.createBufferSource();
    source.buffer = buffer;
    source.loop = true;
    source.connect(trackGain);
    const startOffset = Math.min(offsetSec % Math.max(buffer.duration, 0.001), buffer.duration - 0.001);
    source.start(baseTime + delaySec, Math.max(0, startOffset));
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
        startBgMusic();
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
      startBgMusic(isPlaying ? playbackOffsetSec : 0);
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