// Pronunciation scoring with Azure Speech (word by word, sound by sound) and a reference voice.
// Needs a logged-in account: the token comes from our Cloudflare server, never a key in the page.
const SDK_URL = "https://cdn.jsdelivr.net/npm/microsoft-cognitiveservices-speech-sdk@1.52.0/distrib/browser/microsoft.cognitiveservices.speech.sdk.bundle-min.js";
const VOICE = "en-GB-RyanNeural"; // Alex: an English student with clear, international English

let sdkPromise = null;
function loadSdk() {
  if (window.SpeechSDK) return Promise.resolve(window.SpeechSDK);
  sdkPromise ??= new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = SDK_URL; s.async = true;
    s.onload = () => window.SpeechSDK ? res(window.SpeechSDK) : rej(new Error("sdk"));
    s.onerror = () => { sdkPromise = null; rej(new Error("sdk")); };
    document.head.appendChild(s);
  });
  return sdkPromise;
}

let cached = null; // {token, region, until}
async function token() {
  if (cached && Date.now() < cached.until) return cached;
  if (!window.Account?.signedIn()) throw Object.assign(new Error("login"), { code: "login" });
  const t = await window.Account.speechToken();
  cached = { token: t.token, region: t.region, until: Date.now() + (t.expiresIn || 540) * 1000 };
  return cached;
}

async function config() {
  const [SDK, t] = await Promise.all([loadSdk(), token()]);
  return { SDK, cfg: SDK.SpeechConfig.fromAuthorizationToken(t.token, t.region) };
}

// Listen once to the learner reading `reference`; resolves with scores per sentence, word and sound.
// Live microphone level (0..1) for the "listening" waves. Returns a stop function.
async function meter(onLevel) {
  let stream, ctx, raf;
  try {
    stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    ctx = new (window.AudioContext || window.webkitAudioContext)(); await ctx.resume();
    const an = ctx.createAnalyser(); an.fftSize = 512; ctx.createMediaStreamSource(stream).connect(an);
    const buf = new Uint8Array(an.fftSize);
    const tick = () => { an.getByteTimeDomainData(buf); let sum = 0; for (const v of buf) sum += (v - 128) ** 2; onLevel(Math.min(1, Math.sqrt(sum / buf.length) / 40)); raf = requestAnimationFrame(tick); };
    tick();
  } catch { onLevel(-1); } // no level available: the waves just pulse
  return () => { cancelAnimationFrame(raf); stream?.getTracks().forEach(t => t.stop()); ctx?.close().catch(() => {}); };
}
// Run one recognition; control.stop() (set here) ends it early and rejects with code "stopped".
function stoppable(rec, control) {
  return new Promise((res, rej) => {
    let over = false;
    if (control) control.stop = () => { if (over) return; over = true; try { rec.close(); } catch {} rej(Object.assign(new Error("stopped"), { code: "stopped" })); };
    rec.recognizeOnceAsync(r => { if (over) return; over = true; rec.close(); res(r); }, e => { if (over) return; over = true; try { rec.close(); } catch {} rej(e); });
  });
}
// Wire the phases of a recognition to the page: listening, hearing speech, checking.
function phases(rec, { onListening, onSpeech, onChecking } = {}) {
  rec.sessionStarted = () => onListening?.();
  rec.speechStartDetected = () => onSpeech?.();
  rec.speechEndDetected = () => onChecking?.();
}

async function assess(reference, { onListening, onSpeech, onChecking, control } = {}) {
  const { SDK, cfg } = await config();
  cfg.speechRecognitionLanguage = "en-US";
  const audio = SDK.AudioConfig.fromDefaultMicrophoneInput();
  const pa = new SDK.PronunciationAssessmentConfig(reference, SDK.PronunciationAssessmentGradingSystem.HundredMark, SDK.PronunciationAssessmentGranularity.Phoneme, true);
  pa.phonemeAlphabet = "IPA";
  try { pa.enableProsodyAssessment = true; } catch {}
  const rec = new SDK.SpeechRecognizer(cfg, audio);
  pa.applyTo(rec);
  phases(rec, { onListening, onSpeech, onChecking });
  const result = await stoppable(rec, control);
  if (result.reason === SDK.ResultReason.Canceled) {
    const det = SDK.CancellationDetails.fromResult(result);
    throw Object.assign(new Error(det.errorDetails || "canceled"), { code: /permission|NotAllowed|microphone/i.test(det.errorDetails || "") ? "mic" : "canceled" });
  }
  if (result.reason !== SDK.ResultReason.RecognizedSpeech) throw Object.assign(new Error("nomatch"), { code: "nomatch" });
  const json = JSON.parse(result.properties.getProperty(SDK.PropertyId.SpeechServiceResponse_JsonResult));
  const best = json.NBest?.[0] || {};
  const s = best.PronunciationAssessment || {};
  return {
    heard: result.text,
    pron: Math.round(s.PronScore ?? 0), accuracy: Math.round(s.AccuracyScore ?? 0),
    fluency: Math.round(s.FluencyScore ?? 0), completeness: Math.round(s.CompletenessScore ?? 0),
    prosody: s.ProsodyScore != null ? Math.round(s.ProsodyScore) : null,
    words: (best.Words || []).map(w => ({
      word: w.Word, score: Math.round(w.PronunciationAssessment?.AccuracyScore ?? 0),
      error: w.PronunciationAssessment?.ErrorType || "None",
      phonemes: (w.Phonemes || []).map(p => ({ ipa: p.Phoneme, score: Math.round(p.PronunciationAssessment?.AccuracyScore ?? 0) })),
    })),
  };
}

// Free speech to text (no reference sentence): for "say it in English" and "remember the word".
async function listen({ onListening, onSpeech, onChecking, control, lang = "en-US" } = {}) {
  const { SDK, cfg } = await config();
  cfg.speechRecognitionLanguage = lang;
  const rec = new SDK.SpeechRecognizer(cfg, SDK.AudioConfig.fromDefaultMicrophoneInput());
  phases(rec, { onListening, onSpeech, onChecking });
  const result = await stoppable(rec, control);
  if (result.reason === SDK.ResultReason.Canceled) {
    const det = SDK.CancellationDetails.fromResult(result);
    throw Object.assign(new Error(det.errorDetails || "canceled"), { code: /permission|NotAllowed|microphone/i.test(det.errorDetails || "") ? "mic" : "canceled" });
  }
  if (result.reason !== SDK.ResultReason.RecognizedSpeech) throw Object.assign(new Error("nomatch"), { code: "nomatch" });
  return result.text;
}

// The sentence read by a natural voice; cached so replays are free.
const voiceCache = new Map();
// One shared element, unlocked on the first tap, so Safari lets it play after the voice is fetched.
const voicePlayer = new Audio();
const SILENCE = "data:audio/mpeg;base64,SUQzBAAAAAAAI1RTU0UAAAAPAAADTGF2ZjYwLjE2LjEwMAAAAAAAAAAAAAAA//OEwAAAAAAAAAAAAEluZm8AAAAPAAAABQAAAqAAbW1tbW1tbW1tbW1tbW1tbW1tbZKSkpKSkpKSkpKSkpKSkpKSkpKStra2tra2tra2tra2tra2tra2trbb29vb29vb29vb29vb29vb29vb2///////////////////////////AAAAAExhdmM2MC4zMQAAAAAAAAAAAAAAACQEUAAAAAAAAAKgvT9A1gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//NExAAAAANIAAAAAExBTUUzLjEwMFVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVTEFNRTMu//NExFMAAANIAAAAADEwMFVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVTEFNRTMu//NExKYAAANIAAAAADEwMFVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NExKwAAANIAAAAAFVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NExKwAAANIAAAAAFVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV";
let unlocked = false, actx = null;
function unlock() {
  if (unlocked) return; unlocked = true;
  voicePlayer.src = SILENCE; voicePlayer.play().catch(() => {});
  try { actx = new (window.AudioContext || window.webkitAudioContext)(); actx.resume(); } catch {}
}

// One synthesis, cached: mp3 bytes + when each word is said. rate < 1 = slower (SSML).
const synthCache = new Map();
const xmlEsc = s => s.replace(/[<>&'"]/g, c => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;" })[c]);
function synth(text, voice = VOICE, rate = 1, ph = null) {
  const key = voice + "|" + rate + "|" + text + "|" + JSON.stringify(ph || {});
  if (!synthCache.has(key)) synthCache.set(key, (async () => {
    const { SDK, cfg } = await config();
    cfg.speechSynthesisVoiceName = voice;
    cfg.speechSynthesisOutputFormat = SDK.SpeechSynthesisOutputFormat.Audio24Khz48KBitRateMonoMp3;
    const s = new SDK.SpeechSynthesizer(cfg, null);
    const words = [];
    s.wordBoundary = (_, e) => {
      if (e.boundaryType !== undefined && String(e.boundaryType).toLowerCase().includes("punctuation")) return;
      const start = e.audioOffset / 1e7, dur = e.duration ? (e.duration > 1e4 ? e.duration / 1e7 : e.duration / 1000) : .25;
      words.push({ text: e.text, start, end: start + Math.max(.08, dur) });
    };
    // ph = {word: "IPA"}: say that word exactly this way (e.g. stress on the last syllable)
    let body = xmlEsc(text);
    for (const [w, ipa] of Object.entries(ph || {})) body = body.replace(new RegExp("\\b(" + w + ")\\b", "i"), `<phoneme alphabet="ipa" ph="${xmlEsc(ipa)}">$1</phoneme>`);
    const ssml = `<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="en-US"><voice name="${voice}"><prosody rate="${Math.round((rate - 1) * 100)}%">${body}</prosody></voice></speak>`;
    const result = await new Promise((res, rej) => rate === 1 && !ph ? s.speakTextAsync(text, res, rej) : s.speakSsmlAsync(ssml, res, rej));
    s.close();
    if (result.reason !== SDK.ResultReason.SynthesizingAudioCompleted) { synthCache.delete(key); throw new Error("tts"); }
    return { data: result.audioData, url: URL.createObjectURL(new Blob([result.audioData], { type: "audio/mpeg" })), words };
  })().catch(e => { synthCache.delete(key); throw e; }));
  return synthCache.get(key);
}

async function speak(text, voice = VOICE, rate = 1) {
  unlock(); // must run synchronously inside the tap
  const hit = await synth(text, voice, rate);
  voicePlayer.pause(); voicePlayer.src = hit.url; voicePlayer.onended = null;
  await voicePlayer.play();
  return voicePlayer;
}
// Like speak(), but also returns when each word is said (seconds), for lip-sync and captions.
async function speakTimed(text, voice = VOICE, rate = 1, ph = null) {
  unlock();
  const hit = await synth(text, voice, rate, ph);
  voicePlayer.pause(); voicePlayer.src = hit.url; voicePlayer.onended = null;
  await voicePlayer.play();
  return { player: voicePlayer, words: hit.words };
}

// Several voices saying the same line together (a crowd): mixed with Web Audio, slightly out of step.
// Returns a player-like object, so finished() and lip-sync work the same way.
let chorusNow = null;
async function chorus(text, voices) {
  unlock();
  const hits = await Promise.all(voices.map(v => synth(text, v)));
  if (!actx) { return { player: await speak(text, voices[0]), words: hits[0].words }; }
  await actx.resume();
  const bufs = await Promise.all(hits.map(h => actx.decodeAudioData(h.data.slice(0))));
  voicePlayer.pause(); chorusNow?.stop();
  const t0 = actx.currentTime + .05, gain = actx.createGain(); gain.gain.value = Math.min(1, 1.6 / Math.sqrt(bufs.length)); gain.connect(actx.destination);
  const srcs = bufs.map((b, i) => { const s = actx.createBufferSource(); s.buffer = b; s.connect(gain); s.start(t0 + i * .045 + Math.random() * .03); return s; });
  const dur = Math.max(...bufs.map(b => b.duration)) + bufs.length * .045;
  const p = {
    ended: false, paused: false, onended: null, onpause: null,
    get currentTime() { return Math.max(0, actx.currentTime - t0); },
    stop() { if (p.paused) return; srcs.forEach(s => { try { s.stop(); } catch {} }); p.paused = true; p.onpause?.(); },
  };
  setTimeout(() => { if (!p.paused) { p.ended = true; p.paused = true; p.onended?.(); } }, dur * 1000 + 60);
  chorusNow = p;
  return { player: p, words: hits[0].words };
}

// Resolves when the current line has finished (or was stopped).
const finished = p => new Promise(res => { if (p.ended || p.paused) return res(); p.onended = res; p.onpause = res; });
function stop() { voicePlayer.pause(); chorusNow?.stop(); }

// Tips in Italian for the classic traps; anything else gets a generic hint from the sounds.
const TIPS = {
  hour: "la h non si pronuncia: «àuer», come «our».",
  wrong: "la w è muta: si dice «rong».",
  waited: "la -ed finale qui suona «-id»: «uéitid».",
  gate: "vocale lunga: «gheit», come in «hey». Non «ghet».",
  boarding: "«bòording»: la r quasi non si sente, la o è lunga.",
  will: "breve e morbida: «uil», la w come una u veloce.",
  start: "«staart», la a lunga e aperta.",
  minutes: "«mìnits»: la seconda i quasi sparisce.",
  where: "«uèr»: la h non si sente.",
  excuse: "«ikschiùs»: l'accento cade su -cuse.",
  twenty: "«tuènti»: le due t sono leggere.",
  twentyseven: "«tuènti-sèvn»: due accenti, su tuèn- e su sè-.",
  seven: "«sèvn»: la seconda e quasi non si sente.",
  the: "lingua tra i denti: è il suono «th», né «d» né «z».",
  for: "debole e veloce, quasi «fer».",
  nine: "«nain», con il dittongo «ai».",
  ten: "«ten», breve.",
  yesterday: "«ièsterdei»: l'accento sulla prima sillaba.",
};
function tipFor(w) {
  const key = w.word.toLowerCase().replace(/[^a-z']/g, "");
  if (w.error === "Omission") return `«${w.word}»: non l'ho sentita. Leggila senza saltarla.`;
  if (TIPS[key]) return `«${w.word}»: ${TIPS[key]}`;
  const worst = [...w.phonemes].sort((a, b) => a.score - b.score)[0];
  return worst && worst.score < 60 ? `«${w.word}»: il suono /${worst.ipa}/ non era chiaro. Ascoltala e ripeti piano.` : `«${w.word}»: ascoltala e ripeti piano.`;
}

window.Pronuncia = { assess, listen, meter, speak, speakTimed, chorus, finished, stop, tipFor, available: () => !!window.Account?.signedIn() };
document.dispatchEvent(new Event("pronuncia:ready"));
