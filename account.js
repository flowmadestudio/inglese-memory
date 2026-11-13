// Login (Firebase) + saving progress to the account (Cloudflare Worker).
// The Firebase config below is public by design: it only names the project.
import { initializeApp } from "https://www.gstatic.com/firebasejs/11.0.2/firebase-app.js";
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  sendPasswordResetEmail, signInWithPopup, GoogleAuthProvider, signOut, setPersistence, browserLocalPersistence,
} from "https://www.gstatic.com/firebasejs/11.0.2/firebase-auth.js";

const firebaseConfig = {
  apiKey: "AIzaSyD0NxLkQyY1TQhV68U11aR0bnJu5hoAeUE",
  authDomain: "inglese-memory.firebaseapp.com",
  projectId: "inglese-memory",
  storageBucket: "inglese-memory.firebasestorage.app",
  messagingSenderId: "42854143095",
  appId: "1:42854143095:web:77bbe1a2a6f7d73cc3e272",
};
const API = "https://inglese-memory-api.inglese-memory.workers.dev";

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
auth.languageCode = "it";
setPersistence(auth, browserLocalPersistence).catch(() => {});

// ---------- UI ----------
const css = `
.acct-veil { position: fixed; inset: 0; z-index: 50; background: var(--bg); display: grid; place-items: center; overflow-y: auto;
  padding: max(20px, env(safe-area-inset-top, 0px)) max(16px, env(safe-area-inset-right, 0px)) max(20px, env(safe-area-inset-bottom, 0px)) max(16px, env(safe-area-inset-left, 0px)); }
.acct-box { width: 100%; max-width: 400px; background: var(--card); border: 1px solid var(--line); border-radius: 20px; padding: 26px 22px; display: grid; gap: 14px; }
.acct-box img { width: 96px; height: auto; justify-self: center; }
.acct-box h1 { text-align: center; font-size: 1.7rem; }
.acct-box p { text-align: center; justify-self: center; }
.acct-box label { display: grid; gap: 4px; font-weight: 700; font-size: .92rem; }
.acct-box input { width: 100%; font: 600 1.05rem var(--font-body); padding: .7em .9em; border-radius: 12px; border: 2px solid var(--line); background: var(--bg); color: var(--fg); }
.acct-box button { width: 100%; min-height: 48px; }
.acct-or { display: flex; align-items: center; gap: 10px; color: var(--muted); font-size: .85rem; font-weight: 700; }
.acct-or::before, .acct-or::after { content: ""; flex: 1; height: 1px; background: var(--line); }
.acct-links { display: flex; justify-content: space-between; flex-wrap: wrap; gap: 6px; }
.acct-links button { width: auto; min-height: 40px; border: 0; padding: .3em 0; color: var(--key); background: none; font-size: .92rem; }
.acct-msg { border-radius: 12px; padding: 10px 12px; font-weight: 600; font-size: .95rem; }
.acct-msg.no { background: var(--warn-bg); color: var(--warn); }
.acct-msg.ok { background: var(--ok-bg); color: var(--ok); }
.acct-chip { display: flex; align-items: center; gap: 8px; font-size: .88rem; color: var(--muted); font-weight: 700; min-width: 0; }
.acct-chip span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 22ch; }
.acct-chip button { font-size: .85rem; padding: .35em .8em; }
.acct-sync { font-size: .8rem; }
`;
document.head.insertAdjacentHTML("beforeend", `<style>${css}</style>`);

const veil = document.createElement("div");
veil.className = "acct-veil";
veil.innerHTML = `
  <form class="acct-box" novalidate>
    <img src="head.png" alt="Alex, il pupazzo di feltro">
    <h1 id="acctTitle">Entra</h1>
    <p class="muted" id="acctSub">Così i tuoi progressi restano tuoi, sul telefono e sul computer.</p>
    <div class="acct-msg" id="acctMsg" role="status" hidden></div>
    <button type="button" id="acctGoogle">Continua con Google</button>
    <div class="acct-or">oppure con l'email</div>
    <label>Email <input type="email" id="acctEmail" autocomplete="email" required></label>
    <label>Password <input type="password" id="acctPass" autocomplete="current-password" minlength="6" required></label>
    <button type="submit" class="primary" id="acctSubmit">Entra</button>
    <div class="acct-links">
      <button type="button" id="acctMode">Non hai un account? Crealo</button>
      <button type="button" id="acctForgot">Password dimenticata?</button>
    </div>
  </form>`;
document.body.appendChild(veil);

const $ = id => document.getElementById(id);
let mode = "login";
function setMode(m) {
  mode = m;
  $("acctTitle").textContent = m === "login" ? "Entra" : "Crea il tuo account";
  $("acctSubmit").textContent = m === "login" ? "Entra" : "Crea account";
  $("acctMode").textContent = m === "login" ? "Non hai un account? Crealo" : "Hai già un account? Entra";
  $("acctPass").autocomplete = m === "login" ? "current-password" : "new-password";
  $("acctForgot").hidden = m !== "login";
  msg("");
}
function msg(text, kind = "no") { const el = $("acctMsg"); el.hidden = !text; el.textContent = text; el.className = "acct-msg " + kind; }
function busy(on) { veil.querySelectorAll("button, input").forEach(el => el.disabled = on); }

const ERRORS = {
  "auth/invalid-email": "Questa email non sembra valida.",
  "auth/missing-password": "Scrivi la password.",
  "auth/weak-password": "La password deve avere almeno 6 caratteri.",
  "auth/email-already-in-use": "Esiste già un account con questa email. Prova a entrare.",
  "auth/invalid-credential": "Email o password non corrette.",
  "auth/wrong-password": "Email o password non corrette.",
  "auth/user-not-found": "Non c'è un account con questa email. Crealo.",
  "auth/too-many-requests": "Troppi tentativi. Aspetta qualche minuto e riprova.",
  "auth/popup-closed-by-user": "Hai chiuso la finestra di Google prima di finire.",
  "auth/popup-blocked": "Il browser ha bloccato la finestra di Google: permetti i popup e riprova.",
  "auth/network-request-failed": "Sembra che manchi la connessione. Riprova.",
  "auth/unauthorized-domain": "Questo indirizzo non è autorizzato per il login.",
};
const say = e => ERRORS[e?.code] || "Qualcosa non ha funzionato. Riprova.";

$("acctMode").addEventListener("click", () => setMode(mode === "login" ? "signup" : "login"));
veil.querySelector("form").addEventListener("submit", async e => {
  e.preventDefault();
  const email = $("acctEmail").value.trim(), pass = $("acctPass").value;
  if (!email) return msg("Scrivi la tua email.");
  busy(true); msg("");
  try {
    if (mode === "login") await signInWithEmailAndPassword(auth, email, pass);
    else await createUserWithEmailAndPassword(auth, email, pass);
  } catch (err) { msg(say(err)); }
  busy(false);
});
$("acctGoogle").addEventListener("click", async () => {
  busy(true); msg("");
  try { await signInWithPopup(auth, new GoogleAuthProvider()); } catch (err) { msg(say(err)); }
  busy(false);
});
$("acctForgot").addEventListener("click", async () => {
  const email = $("acctEmail").value.trim();
  if (!email) return msg("Scrivi prima la tua email, poi premi di nuovo «Password dimenticata?».");
  busy(true);
  try { await sendPasswordResetEmail(auth, email); msg("Ti ho mandato un'email per scegliere una nuova password. Guarda anche nello spam.", "ok"); }
  catch (err) { msg(say(err)); }
  busy(false);
});

// Account chip in the top bar
const chip = document.createElement("div");
chip.className = "acct-chip";
chip.hidden = true;
chip.innerHTML = `<span id="acctWho"></span><span class="acct-sync" id="acctSync"></span><button type="button" id="acctOut">Esci</button>`;
document.querySelector(".top")?.appendChild(chip);
$("acctOut").addEventListener("click", () => signOut(auth));

// ---------- Sync ----------
const KEY = word => word + "-progress";
const readLocal = word => { try { return JSON.parse(localStorage.getItem(KEY(word)) || "{}"); } catch { return {}; } };
const writeLocal = (word, data) => { try { localStorage.setItem(KEY(word), JSON.stringify(data)); } catch {} };
const syncState = t => { $("acctSync").textContent = t; };

async function api(path, opts = {}) {
  const user = auth.currentUser;
  if (!user || API.startsWith("__")) throw new Error("offline");
  const token = await user.getIdToken();
  const res = await fetch(API + path, { ...opts, headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json", ...(opts.headers || {}) } });
  if (!res.ok) throw new Error("api " + res.status);
  return res.json();
}

const pending = new Map();
async function push(word, data) {
  pending.set(word, data);
  try { await api("/progress/" + encodeURIComponent(word), { method: "PUT", body: JSON.stringify({ data }) }); pending.delete(word); syncState("· salvato"); }
  catch { syncState("· salvato solo qui"); }
}
window.addEventListener("online", () => pending.forEach((d, w) => push(w, d)));

// The lesson page calls this whenever a word's progress changes.
window.Account = {
  save(word, data) { writeLocal(word, data); if (auth.currentUser) push(word, data); },
  load: readLocal,
};

// Each account gets its own copy in this browser: switching account must not mix them.
const OWNER = "account-owner";
async function pullAll(user) {
  let prevOwner = null; try { prevOwner = localStorage.getItem(OWNER); } catch {}
  if (prevOwner && prevOwner !== user.uid) {
    try { Object.keys(localStorage).filter(k => k.endsWith("-progress")).forEach(k => localStorage.removeItem(k)); } catch {}
  }
  try { localStorage.setItem(OWNER, user.uid); } catch {}
  try {
    const me = await api("/me");
    for (const [word, data] of Object.entries(me.progress)) writeLocal(word, data);
    // Progress made here before logging in, and not yet on the account: send it up.
    if (!prevOwner) {
      for (const k of Object.keys(localStorage).filter(k => k.endsWith("-progress"))) {
        const word = k.slice(0, -"-progress".length);
        if (!me.progress[word]) push(word, readLocal(word));
      }
    }
    syncState("· sincronizzato");
  } catch { syncState("· solo su questo dispositivo"); }
  document.dispatchEvent(new CustomEvent("account:ready", { detail: { uid: user.uid } }));
}

onAuthStateChanged(auth, async user => {
  if (user) {
    veil.hidden = true; chip.hidden = false;
    $("acctWho").textContent = user.displayName || user.email || "Account";
    await pullAll(user);
  } else {
    veil.hidden = false; chip.hidden = true; setMode("login");
    $("acctPass").value = "";
  }
});
