// Inglese Memory API: checks the Firebase login, keeps each account's progress in D1.
// Secrets never live in the site: the browser only ever sends its Firebase ID token.

const PROJECT_ID = "inglese-memory";
const ALLOWED_ORIGINS = ["https://flowmadestudio.github.io", "http://localhost:8000", "http://127.0.0.1:8000"];
const JWKS_URL = "https://www.googleapis.com/service_accounts/v1/jwk/securetoken@system.gserviceaccount.com";
const MAX_BODY = 32 * 1024;
const WORD_RE = /^[a-z][a-z' -]{0,39}$/;
// One token covers ~9 minutes of reading; 60 a day is far more than a real session needs.
const SPEECH_TOKENS_PER_DAY = 60;

const PRON_PER_DAY = 300;              // pronunciation checks per account per day
const MAX_AUDIO = 1.5 * 1024 * 1024;   // ~45 s of 16 kHz mono WAV

// Pronunciation assessment through Azure's REST API: the browser sends a WAV, we return the scores.
async function azurePronounce(env, wav, reference) {
  const cfg = { ReferenceText: reference, GradingSystem: "HundredMark", Granularity: "Phoneme", Dimension: "Comprehensive", EnableMiscue: true, EnableProsodyAssessment: true, PhonemeAlphabet: "IPA" };
  const header = btoa(unescape(encodeURIComponent(JSON.stringify(cfg))));
  const res = await fetch(`https://${env.AZURE_SPEECH_REGION}.stt.speech.microsoft.com/speech/recognition/conversation/cognitiveservices/v1?language=en-US&format=detailed`, {
    method: "POST",
    headers: { "Ocp-Apim-Subscription-Key": env.AZURE_SPEECH_KEY, "Content-Type": "audio/wav; codecs=audio/pcm; samplerate=16000", "Pronunciation-Assessment": header, "Accept": "application/json" },
    body: wav,
  });
  const text = await res.text();
  if (!res.ok) { console.error("azure pron", res.status, text.slice(0, 300)); return { error: "azure_" + res.status }; }
  return JSON.parse(text);
}
// Only for the self-test: a sentence spoken by an Azure voice, as WAV.
async function azureTtsWav(env, text) {
  const ssml = `<speak version="1.0" xml:lang="en-US"><voice name="en-GB-RyanNeural">${text}</voice></speak>`;
  const res = await fetch(`https://${env.AZURE_SPEECH_REGION}.tts.speech.microsoft.com/cognitiveservices/v1`, {
    method: "POST", headers: { "Ocp-Apim-Subscription-Key": env.AZURE_SPEECH_KEY, "Content-Type": "application/ssml+xml", "X-Microsoft-OutputFormat": "riff-16khz-16bit-mono-pcm", "User-Agent": "inglese-memory" }, body: ssml,
  });
  if (!res.ok) throw new Error("tts " + res.status);
  return res.arrayBuffer();
}

async function azureToken(env) {
  if (!env.AZURE_SPEECH_KEY || !env.AZURE_SPEECH_REGION) return null;
  const res = await fetch(`https://${env.AZURE_SPEECH_REGION}.api.cognitive.microsoft.com/sts/v1.0/issueToken`, {
    method: "POST", headers: { "Ocp-Apim-Subscription-Key": env.AZURE_SPEECH_KEY, "Content-Length": "0" },
  });
  if (!res.ok) { console.error("azure token", res.status); return null; }
  return res.text();
}

export default {
  async fetch(request, env, ctx) {
    const origin = request.headers.get("Origin") || "";
    const cors = corsHeaders(origin);
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });

    try {
      const url = new URL(request.url);
      if (url.pathname === "/health") {
        // ?speech=1 checks that the Azure key still works, without exposing anything
        if (url.searchParams.get("speech") === "1") return json({ ok: true, speech: !!(await azureToken(env)) }, 200, cors);
        // ?pa=<DEBUG_KEY>: full self-test of the pronunciation path (Azure voice → Azure assessment)
        if (env.DEBUG_KEY && url.searchParams.get("pa") === env.DEBUG_KEY && url.searchParams.get("tts")) {
          return new Response(await azureTtsWav(env, url.searchParams.get("tts").slice(0, 200)), { headers: { "Content-Type": "audio/wav", ...cors } });
        }
        if (env.DEBUG_KEY && url.searchParams.get("pa") === env.DEBUG_KEY && request.method === "POST") {
          return json(await azurePronounce(env, await request.arrayBuffer(), (url.searchParams.get("ref") || "").slice(0, 300)), 200, cors);
        }
        if (env.DEBUG_KEY && url.searchParams.get("pa") === env.DEBUG_KEY) {
          const ref = "It was a unique experience, and I'll never forget it.";
          const t0 = Date.now(); const wav = await azureTtsWav(env, ref); const t1 = Date.now();
          const r = await azurePronounce(env, wav, ref);
          const best = r.NBest?.[0] || {};
          return json({ ok: !r.error, error: r.error, status: r.RecognitionStatus, heard: r.DisplayText, ttsMs: t1 - t0, assessMs: Date.now() - t1, wavBytes: wav.byteLength,
            scores: best.PronunciationAssessment || { AccuracyScore: best.AccuracyScore, FluencyScore: best.FluencyScore, CompletenessScore: best.CompletenessScore, PronScore: best.PronScore },
            firstWord: best.Words?.[0] }, 200, cors);
        }
        return json({ ok: true }, 200, cors);
      }

      const user = await authenticate(request, ctx);
      if (!user) return json({ error: "login_required" }, 401, cors);
      await touchUser(env, user);

      // GET /me → everything this account has saved
      if (url.pathname === "/me" && request.method === "GET") {
        const progress = await env.DB.prepare("SELECT word, data, updated_at FROM progress WHERE uid = ?").bind(user.uid).all();
        const queue = await env.DB.prepare("SELECT word FROM queue WHERE uid = ? ORDER BY position").bind(user.uid).all();
        const out = {};
        for (const r of progress.results) out[r.word] = { ...safeParse(r.data), updatedAt: r.updated_at };
        return json({ uid: user.uid, email: user.email, progress: out, queue: queue.results.map(r => r.word) }, 200, cors);
      }

      // PUT /progress/<word> {data} → save one word's progress
      const m = url.pathname.match(/^\/progress\/([^/]+)$/);
      if (m && request.method === "PUT") {
        const word = decodeURIComponent(m[1]).toLowerCase();
        if (!WORD_RE.test(word)) return json({ error: "bad_word" }, 400, cors);
        const body = await readJson(request);
        if (!body || typeof body.data !== "object") return json({ error: "bad_body" }, 400, cors);
        const now = Date.now();
        await env.DB.prepare(
          "INSERT INTO progress (uid, word, data, updated_at) VALUES (?, ?, ?, ?) " +
          "ON CONFLICT(uid, word) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at"
        ).bind(user.uid, word, JSON.stringify(body.data), now).run();
        return json({ ok: true, updatedAt: now }, 200, cors);
      }

      // PUT /queue {words: [...]} → replace the account's word queue
      if (url.pathname === "/queue" && request.method === "PUT") {
        const body = await readJson(request);
        const words = Array.isArray(body?.words) ? [...new Set(body.words.map(w => String(w).toLowerCase()))] : null;
        if (!words || words.length > 3000 || !words.every(w => WORD_RE.test(w))) return json({ error: "bad_body" }, 400, cors);
        const stmts = [env.DB.prepare("DELETE FROM queue WHERE uid = ?").bind(user.uid)];
        words.forEach((w, i) => stmts.push(env.DB.prepare("INSERT INTO queue (uid, position, word) VALUES (?, ?, ?)").bind(user.uid, i, w)));
        await env.DB.batch(stmts);
        return json({ ok: true }, 200, cors);
      }

      // GET /speech-token → a 10-minute Azure Speech token; the key itself never leaves the server
      if (url.pathname === "/speech-token" && request.method === "GET") {
        const day = new Date().toISOString().slice(0, 10);
        const row = await env.DB.prepare(
          "INSERT INTO speech_usage (uid, day, count) VALUES (?, ?, 1) " +
          "ON CONFLICT(uid, day) DO UPDATE SET count = count + 1 RETURNING count"
        ).bind(user.uid, day).first();
        if (row.count > SPEECH_TOKENS_PER_DAY) return json({ error: "daily_limit" }, 429, cors);
        const token = await azureToken(env);
        if (!token) return json({ error: "speech_unavailable" }, 502, cors);
        return json({ token, region: env.AZURE_SPEECH_REGION, expiresIn: 540 }, 200, { ...cors, "Cache-Control": "no-store" });
      }

      // POST /pronunciation?ref=<sentence>  body: 16 kHz mono WAV → Azure's scores
      if (url.pathname === "/pronunciation" && request.method === "POST") {
        const ref = (url.searchParams.get("ref") || "").slice(0, 300);
        if (!ref) return json({ error: "bad_ref" }, 400, cors);
        const wav = await request.arrayBuffer();
        if (wav.byteLength < 2000 || wav.byteLength > MAX_AUDIO) return json({ error: "bad_audio" }, 400, cors);
        const day = "p" + new Date().toISOString().slice(0, 10);
        const row = await env.DB.prepare("INSERT INTO speech_usage (uid, day, count) VALUES (?, ?, 1) ON CONFLICT(uid, day) DO UPDATE SET count = count + 1 RETURNING count").bind(user.uid, day).first();
        if (row.count > PRON_PER_DAY) return json({ error: "daily_limit" }, 429, cors);
        const r = await azurePronounce(env, wav, ref);
        if (r.error) return json(r, 502, cors);
        return json(r, 200, { ...cors, "Cache-Control": "no-store" });
      }

      return json({ error: "not_found" }, 404, cors);
    } catch (err) {
      console.error(err);
      return json({ error: "server_error" }, 500, cors);
    }
  },
};

function corsHeaders(origin) {
  const h = { "Access-Control-Allow-Methods": "GET, PUT, POST, OPTIONS", "Access-Control-Allow-Headers": "Authorization, Content-Type", "Access-Control-Max-Age": "86400", "Vary": "Origin" };
  if (ALLOWED_ORIGINS.includes(origin)) h["Access-Control-Allow-Origin"] = origin;
  return h;
}

function json(obj, status, extra) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json; charset=utf-8", ...extra } });
}

function safeParse(s) { try { return JSON.parse(s); } catch { return {}; } }

async function readJson(request) {
  const text = await request.text();
  if (text.length > MAX_BODY) return null;
  return safeParse(text);
}

async function touchUser(env, user) {
  const now = Date.now();
  await env.DB.prepare(
    "INSERT INTO users (uid, email, created_at, last_seen) VALUES (?, ?, ?, ?) " +
    "ON CONFLICT(uid) DO UPDATE SET email = excluded.email, last_seen = excluded.last_seen"
  ).bind(user.uid, user.email || null, now, now).run();
}

// ---- Firebase ID token check (RS256, Google's public keys) ----
let jwksCache = { keys: null, until: 0 };

async function getKeys(ctx) {
  if (jwksCache.keys && Date.now() < jwksCache.until) return jwksCache.keys;
  const res = await fetch(JWKS_URL, { cf: { cacheTtl: 3600 } });
  if (!res.ok) throw new Error("jwks " + res.status);
  const maxAge = Number((res.headers.get("Cache-Control") || "").match(/max-age=(\d+)/)?.[1] || 3600);
  const { keys } = await res.json();
  jwksCache = { keys, until: Date.now() + maxAge * 1000 };
  return keys;
}

function b64urlToBytes(s) {
  s = s.replace(/-/g, "+").replace(/_/g, "/");
  const bin = atob(s + "===".slice((s.length + 3) % 4));
  return Uint8Array.from(bin, c => c.charCodeAt(0));
}

async function authenticate(request, ctx) {
  const auth = request.headers.get("Authorization") || "";
  const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
  const parts = token.split(".");
  if (parts.length !== 3) return null;
  let header, payload;
  try {
    header = JSON.parse(new TextDecoder().decode(b64urlToBytes(parts[0])));
    payload = JSON.parse(new TextDecoder().decode(b64urlToBytes(parts[1])));
  } catch { return null; }
  if (header.alg !== "RS256" || !header.kid) return null;

  const now = Math.floor(Date.now() / 1000);
  if (payload.aud !== PROJECT_ID) return null;
  if (payload.iss !== "https://securetoken.google.com/" + PROJECT_ID) return null;
  if (typeof payload.sub !== "string" || !payload.sub || payload.sub.length > 128) return null;
  if (!(payload.exp > now - 30) || !(payload.iat <= now + 30) || !(payload.auth_time <= now + 30)) return null;

  const jwk = (await getKeys(ctx)).find(k => k.kid === header.kid);
  if (!jwk) return null;
  const key = await crypto.subtle.importKey("jwk", jwk, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["verify"]);
  const ok = await crypto.subtle.verify("RSASSA-PKCS1-v1_5", key, b64urlToBytes(parts[2]), new TextEncoder().encode(parts[0] + "." + parts[1]));
  if (!ok) return null;
  return { uid: payload.sub, email: payload.email || null };
}
