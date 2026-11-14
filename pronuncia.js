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
async function assess(reference, { onListening } = {}) {
  const { SDK, cfg } = await config();
  cfg.speechRecognitionLanguage = "en-US";
  const audio = SDK.AudioConfig.fromDefaultMicrophoneInput();
  const pa = new SDK.PronunciationAssessmentConfig(reference, SDK.PronunciationAssessmentGradingSystem.HundredMark, SDK.PronunciationAssessmentGranularity.Phoneme, true);
  pa.phonemeAlphabet = "IPA";
  try { pa.enableProsodyAssessment = true; } catch {}
  const rec = new SDK.SpeechRecognizer(cfg, audio);
  pa.applyTo(rec);
  onListening?.();
  const result = await new Promise((res, rej) => rec.recognizeOnceAsync(res, rej));
  rec.close();
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

// The sentence read by a natural voice; cached so replays are free.
const voiceCache = new Map();
let voicePlayer = null;
async function speak(text) {
  let url = voiceCache.get(text);
  if (!url) {
    const { SDK, cfg } = await config();
    cfg.speechSynthesisVoiceName = VOICE;
    cfg.speechSynthesisOutputFormat = SDK.SpeechSynthesisOutputFormat.Audio24Khz48KBitRateMonoMp3;
    const synth = new SDK.SpeechSynthesizer(cfg, null);
    const result = await new Promise((res, rej) => synth.speakTextAsync(text, res, rej));
    synth.close();
    if (result.reason !== SDK.ResultReason.SynthesizingAudioCompleted) throw new Error("tts");
    url = URL.createObjectURL(new Blob([result.audioData], { type: "audio/mpeg" }));
    voiceCache.set(text, url);
  }
  voicePlayer?.pause();
  voicePlayer = new Audio(url);
  await voicePlayer.play();
  return voicePlayer;
}

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

window.Pronuncia = { assess, speak, tipFor, available: () => !!window.Account?.signedIn() };
document.dispatchEvent(new Event("pronuncia:ready"));
