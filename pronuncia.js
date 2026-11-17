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
    if (control) control.stop = (code = "stopped") => { if (over) return; over = true; try { rec.close(); } catch {} rej(Object.assign(new Error(code), { code })); };
    // Never hang: if Azure doesn't answer, give up after 25 s (the page also has its own shorter limit).
    setTimeout(() => control?.stop ? control.stop("timeout") : (!over && (over = true, (() => { try { rec.close(); } catch {} })(), rej(Object.assign(new Error("timeout"), { code: "timeout" })))), 25000);
    rec.recognizeOnceAsync(r => { if (over) return; over = true; rec.close(); res(r); }, e => { if (over) return; over = true; try { rec.close(); } catch {} rej(e); });
  });
}
// Wire the phases of a recognition to the page: listening, hearing speech, checking.
function phases(rec, { onListening, onSpeech, onChecking } = {}) {
  rec.sessionStarted = () => onListening?.();
  rec.speechStartDetected = () => onSpeech?.();
  rec.speechEndDetected = () => onChecking?.();
}

// Record the learner reading `reference`, then our server asks Azure for the scores.
// Recording stops by itself after ~1.2 s of silence (or 15 s at most). control.stop() cancels.
async function assess(reference, { onListening, onSpeech, onChecking, control } = {}) {
  if (!window.Account?.signedIn()) throw Object.assign(new Error("login"), { code: "login" });
  let stream;
  try { stream = await navigator.mediaDevices.getUserMedia({ audio: { channelCount: 1, echoCancellation: true, noiseSuppression: true } }); }
  catch { throw Object.assign(new Error("mic"), { code: "mic" }); }
  const ctx = new (window.AudioContext || window.webkitAudioContext)(); await ctx.resume();
  const src = ctx.createMediaStreamSource(stream), proc = ctx.createScriptProcessor(4096, 1, 1), mute = ctx.createGain();
  mute.gain.value = 0; src.connect(proc); proc.connect(mute); mute.connect(ctx.destination);
  const chunks = [];
  const cleanup = () => { try { proc.disconnect(); src.disconnect(); } catch {} stream.getTracks().forEach(t => t.stop()); ctx.close().catch(() => {}); };
  const recorded = await new Promise((resolve, reject) => {
    const t0 = performance.now(); let floor = null, spoke = false, lastVoice = 0, done = false;
    const finish = (err) => { if (done) return; done = true; cleanup(); err ? reject(err) : resolve(); };
    if (control) control.stop = (code = "stopped") => finish(Object.assign(new Error(code), { code }));
    onListening?.();
    proc.onaudioprocess = e => {
      if (done) return;
      const d = e.inputBuffer.getChannelData(0); chunks.push(new Float32Array(d));
      let sum = 0; for (let i = 0; i < d.length; i++) sum += d[i] * d[i]; const rms = Math.sqrt(sum / d.length);
      const now = performance.now() - t0;
      if (now < 300) { floor = floor == null ? rms : Math.max(floor, rms); return; } // room noise
      const voice = rms > Math.max(.015, (floor || 0) * 2.5);
      if (voice) { lastVoice = now; if (!spoke) { spoke = true; onSpeech?.(); } }
      if (spoke && now - lastVoice > 1200) finish();                                          // finished speaking
      else if (now > 15000) spoke ? finish() : finish(Object.assign(new Error("nomatch"), { code: "nomatch" }));
      else if (!spoke && now > 8000) finish(Object.assign(new Error("nomatch"), { code: "nomatch" })); // never started
    };
  }).then(() => ({ samples: chunks, rate: ctx.sampleRate }));
  onChecking?.();
  const wav = toWav16k(recorded.samples, recorded.rate);
  // While waiting for the server, Annulla / the page's time limit still work.
  const cancelled = new Promise((_, rej) => { if (control) control.stop = (code = "stopped") => rej(Object.assign(new Error(code), { code })); });
  let json;
  try { json = await Promise.race([window.Account.pronounce(wav, reference), cancelled]); }
  catch (e) { throw Object.assign(new Error(e.message), { code: e.code || "server" }); }
  if (json.RecognitionStatus !== "Success" || !json.NBest?.length) throw Object.assign(new Error("nomatch"), { code: "nomatch" });
  const best = json.NBest[0], sc = best.PronunciationAssessment || best;
  const score = o => Math.round(o?.PronunciationAssessment?.AccuracyScore ?? o?.AccuracyScore ?? 0);
  return {
    heard: json.DisplayText || best.Display || "",
    pron: Math.round(sc.PronScore ?? 0), accuracy: Math.round(sc.AccuracyScore ?? 0),
    fluency: Math.round(sc.FluencyScore ?? 0), completeness: Math.round(sc.CompletenessScore ?? 0),
    prosody: sc.ProsodyScore != null ? Math.round(sc.ProsodyScore) : null,
    words: (best.Words || []).map(w => ({ word: w.Word, score: score(w), error: w.PronunciationAssessment?.ErrorType || w.ErrorType || "None",
      phonemes: (w.Phonemes || []).map(p => ({ ipa: p.Phoneme, score: score(p) })) })),
  };
}
// Float32 chunks at the device rate → 16 kHz 16-bit mono WAV (what Azure expects).
function toWav16k(chunks, rate) {
  const len = chunks.reduce((n, c) => n + c.length, 0), all = new Float32Array(len); let o = 0; for (const c of chunks) { all.set(c, o); o += c.length; }
  const ratio = rate / 16000, n = Math.floor(len / ratio), pcm = new Int16Array(n);
  for (let i = 0; i < n; i++) { const a = Math.floor(i * ratio), b = Math.min(len, Math.floor((i + 1) * ratio)); let s = 0; for (let j = a; j < b; j++) s += all[j]; const v = s / Math.max(1, b - a); pcm[i] = Math.max(-1, Math.min(1, v)) * 0x7fff; }
  const buf = new ArrayBuffer(44 + pcm.length * 2), dv = new DataView(buf);
  const str = (off, t) => [...t].forEach((ch, i) => dv.setUint8(off + i, ch.charCodeAt(0)));
  str(0, "RIFF"); dv.setUint32(4, 36 + pcm.length * 2, true); str(8, "WAVE"); str(12, "fmt "); dv.setUint32(16, 16, true); dv.setUint16(20, 1, true); dv.setUint16(22, 1, true);
  dv.setUint32(24, 16000, true); dv.setUint32(28, 32000, true); dv.setUint16(32, 2, true); dv.setUint16(34, 16, true); str(36, "data"); dv.setUint32(40, pcm.length * 2, true);
  new Int16Array(buf, 44).set(pcm);
  return new Blob([buf], { type: "audio/wav" });
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
