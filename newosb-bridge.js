/*
 * Observatoire Prestaterre — V6.15 — client du pont Google Apps Script (partagé).
 *
 * Utilisé par la source OPERATIONS (app.js) et par la source Exigences (requirements.js).
 * Chaque source crée son propre client : les deux projets Apps Script ont des clés distinctes.
 *
 * Sécurité :
 * - la clé d'accès reste uniquement en mémoire (jamais dans l'URL, localStorage ou sessionStorage) ;
 * - la clé n'est envoyée qu'à la fenêtre du pont, après vérification de son origine
 *   (script.google.com ou *.googleusercontent.com) et du jeton de communication ;
 * - les réponses ne sont acceptées que si elles viennent de cette même fenêtre, avec le même jeton ;
 * - aucun transport de secours (GET JSON, JSONP) : un refus reste un refus.
 */
(() => {
  'use strict';

  const TRUSTED_BRIDGE_ORIGIN = /^https:\/\/(?:script\.google\.com|[a-z0-9-]+\.googleusercontent\.com)$/i;
  const KEY_PARAM_NAMES = ['key', 'accesskey', 'access_key', 'cle'];

  function randomToken() {
    try { const a = new Uint8Array(18); crypto.getRandomValues(a); return 'b' + Array.from(a, x => x.toString(16).padStart(2, '0')).join(''); }
    catch { return 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2); }
  }

  // Retire toute clé éventuellement collée dans l'URL (ancien fonctionnement V6.13 « ?key=… »).
  function splitUrl(raw) {
    const text = String(raw || '').trim();
    if (!text) return { url: '', key: '' };
    let u;
    try { u = new URL(text, location.href); } catch { return { url: text, key: '' }; }
    let key = '';
    [...u.searchParams.keys()].forEach(name => {
      if (KEY_PARAM_NAMES.includes(name.toLowerCase())) { key = key || u.searchParams.get(name) || ''; u.searchParams.delete(name); }
    });
    ['mode', 'bridge', 'bridgeToken', 'interactive', 'prefix', '_ts', '_'].forEach(n => u.searchParams.delete(n));
    return { url: u.toString(), key };
  }

  function bridgeUrl(url, params) {
    const u = new URL(String(url || ''), location.href);
    Object.entries(params).forEach(([k, v]) => u.searchParams.set(k, String(v)));
    return u.toString();
  }

  // La fenêtre source d'un message doit être la fenêtre du pont ou l'un de ses cadres internes
  // (Apps Script affiche le contenu HtmlService dans des iframes imbriquées).
  function isWithin(source, host) {
    if (!source || !host) return false;
    let w = source;
    for (let i = 0; i < 6 && w; i++) {
      if (w === host) return true;
      let p = null;
      try { p = w.parent; } catch { p = null; }
      if (!p || p === w) break;
      w = p;
    }
    return false;
  }

  class AuthError extends Error { constructor(message) { super(message); this.name = 'AuthError'; this.authError = true; } }

  function create(name) {
    let key = '';
    // V6.15 : jeton d'un lien de partage, utilisé à la place de la clé (jamais les deux).
    let share = '', visitor = '';
    // Fermeture différée de la fenêtre du pont : annulée si la fenêtre est réutilisée entre-temps.
    let closeTimer = null;
    const st = {
      url: '', host: null, kind: '', token: '', ready: false, readyPromise: null, target: null, targetOrigin: '',
      listener: null, probeTimer: null, pending: new Map(), popupBlocked: false, iframe: null
    };

    function destroy(reason = 'Pont Apps Script réinitialisé.', closeWindow = true) {
      try { if (st.listener) window.removeEventListener('message', st.listener); } catch {}
      try { if (st.probeTimer) clearInterval(st.probeTimer); } catch {}
      try { if (st.target && st.targetOrigin) st.target.postMessage({ type: 'NEWOSB_BRIDGE_CLOSE', token: st.token }, st.targetOrigin); } catch {}
      try { if (closeWindow && st.kind === 'popup' && st.host && !st.host.closed) st.host.close(); } catch {}
      try { if (st.iframe) st.iframe.remove(); } catch {}
      st.pending.forEach(p => { try { clearTimeout(p.timer); p.reject(new Error(reason)); } catch {} });
      st.pending.clear();
      Object.assign(st, { url: '', host: null, kind: '', token: '', ready: false, readyPromise: null, target: null, targetOrigin: '', listener: null, probeTimer: null, popupBlocked: false, iframe: null });
    }

    function probe() {
      const hello = { type: 'NEWOSB_BRIDGE_HELLO', token: st.token };
      // Le HELLO ne contient aucun secret : il sert seulement à réveiller le pont.
      const send = w => { try { w && w.postMessage(hello, '*'); } catch {} };
      send(st.host);
      try { const n = Math.min(8, Number(st.host && st.host.frames && st.host.frames.length) || 0); for (let i = 0; i < n; i++) { send(st.host.frames[i]); try { const m = Math.min(4, Number(st.host.frames[i].frames.length) || 0); for (let j = 0; j < m; j++) send(st.host.frames[i].frames[j]); } catch {} } } catch {}
    }

    function install(host, url, kind, timeoutMs, token) {
      Object.assign(st, { url, host, kind, token, ready: false, target: null, targetOrigin: '' });
      st.readyPromise = new Promise((resolve, reject) => {
        let settled = false;
        const timer = setTimeout(() => {
          if (settled) return; settled = true;
          try { clearInterval(st.probeTimer); } catch {}
          reject(new Error(kind === 'popup' ? 'La fenêtre Google Apps Script ne répond pas. Si Google demande une autorisation, valide-la dans cette fenêtre.' : 'Le pont Apps Script ne répond pas.'));
        }, timeoutMs);
        st.listener = event => {
          const msg = event.data || {};
          if (!msg || typeof msg !== 'object' || msg.token !== token) return;
          if (!TRUSTED_BRIDGE_ORIGIN.test(String(event.origin || ''))) return;
          if (!isWithin(event.source, host)) return;
          if (msg.type === 'NEWOSB_BRIDGE_READY') {
            if (msg.refused) {
              if (!settled) { settled = true; clearTimeout(timer); try { clearInterval(st.probeTimer); } catch {} reject(new Error(msg.error || 'Le pont a refusé ce site (propriété NEWOSB_ALLOWED_ORIGINS).')); }
              return;
            }
            // La première fenêtre valide devient l'unique interlocuteur ; son origine exacte sert de cible.
            if (!st.target) { st.target = event.source; st.targetOrigin = event.origin; }
            if (event.source !== st.target || event.origin !== st.targetOrigin) return;
            st.ready = true;
            try { clearInterval(st.probeTimer); } catch {}
            if (!settled) { settled = true; clearTimeout(timer); resolve(msg); }
            return;
          }
          if (msg.type === 'NEWOSB_BRIDGE_RESPONSE' && msg.id) {
            if (st.target && (event.source !== st.target || event.origin !== st.targetOrigin)) return;
            const p = st.pending.get(msg.id); if (!p) return;
            st.pending.delete(msg.id); clearTimeout(p.timer);
            if (msg.error) p.reject(msg.authError ? new AuthError(msg.error) : new Error(msg.error));
            else if (msg.payload && msg.payload.ok === false) p.reject(msg.payload.authError ? new AuthError(msg.payload.error || 'Accès refusé.') : new Error(msg.payload.error || 'Erreur Apps Script.'));
            else p.resolve(msg.payload);
          }
        };
        window.addEventListener('message', st.listener);
        probe();
        st.probeTimer = setInterval(probe, 500);
      });
      st.readyPromise.catch(() => {});
      return st.readyPromise;
    }

    // À appeler directement dans le gestionnaire du clic (sinon le navigateur bloque la fenêtre).
    function preparePopup(rawUrl) {
      const url = splitUrl(rawUrl).url;
      if (!url) return null;
      if (st.url === url && st.kind === 'popup' && st.host && !st.host.closed && st.readyPromise) { clearTimeout(closeTimer); closeTimer = null; return st.readyPromise; }
      destroy('Nouvelle connexion Apps Script.');
      const token = randomToken();
      const popup = window.open(bridgeUrl(url, { bridge: 1, interactive: 1, bridgeToken: token, _ts: Date.now() }), 'newosb_bridge_' + name, 'popup=yes,width=600,height=470,resizable=yes,scrollbars=yes');
      if (!popup) { st.popupBlocked = true; return null; }
      try { popup.focus(); } catch {}
      return install(popup, url, 'popup', 120000, token);
    }

    function ensure(rawUrl, timeoutMs = 18000) {
      const url = splitUrl(rawUrl).url;
      if (st.url === url && st.readyPromise && (st.kind !== 'popup' || (st.host && !st.host.closed))) return st.readyPromise;
      destroy('Nouvelle tentative Apps Script.');
      const token = randomToken();
      const iframe = document.createElement('iframe');
      iframe.setAttribute('aria-hidden', 'true'); iframe.tabIndex = -1; iframe.title = 'Pont Apps Script';
      iframe.style.cssText = 'position:fixed;width:1px;height:1px;left:-10000px;top:-10000px;border:0;opacity:0;pointer-events:none;';
      iframe.src = bridgeUrl(url, { bridge: 1, bridgeToken: token, _ts: Date.now() });
      (document.body || document.documentElement).appendChild(iframe);
      st.iframe = iframe;
      const p = install(iframe.contentWindow, url, 'iframe', timeoutMs, token);
      p.catch(() => { if (st.iframe === iframe) { try { iframe.remove(); } catch {} } });
      return p;
    }

    async function request(rawUrl, params = {}, timeoutMs = 90000) {
      if (!key && !share) throw new AuthError('Clé d’accès non saisie : la source privée ne peut pas être lue.');
      await ensure(rawUrl, st.kind === 'popup' ? 120000 : 18000);
      if (!st.target || !st.targetOrigin) throw new Error('Pont Apps Script indisponible.');
      return new Promise((resolve, reject) => {
        const id = randomToken();
        const timer = setTimeout(() => { st.pending.delete(id); reject(new Error(`Délai dépassé côté Apps Script (${String(params.mode || params.endpoint || 'requête')}).`)); }, timeoutMs);
        st.pending.set(id, { resolve, reject, timer });
        // Envoi ciblé : uniquement vers l'origine exacte du pont validée au READY, jamais vers « * ».
        st.target.postMessage({ type: 'NEWOSB_BRIDGE_REQUEST', id, params: share ? { ...params, share, ...(visitor ? { visitor } : {}) } : { ...params, key }, token: st.token }, st.targetOrigin);
      });
    }

    function closePopupSoon(delay = 500) {
      if (st.kind !== 'popup') return;
      const host = st.host;
      clearTimeout(closeTimer);
      closeTimer = setTimeout(() => { closeTimer = null; if (st.host === host) destroy('Fenêtre du pont fermée.'); else { try { host && !host.closed && host.close(); } catch {} } }, delay);
    }

    // Après création d'un document (Google Slides), on réutilise la fenêtre ouverte au clic pour l'afficher.
    function navigatePopup(url) {
      if (st.kind !== 'popup' || !st.host || st.host.closed) return false;
      const host = st.host;
      try { if (st.listener) window.removeEventListener('message', st.listener); } catch {}
      try { clearInterval(st.probeTimer); } catch {}
      Object.assign(st, { url: '', host: null, kind: '', ready: false, readyPromise: null, target: null, targetOrigin: '', listener: null });
      try { host.location.href = url; return true; } catch { return false; }
    }

    return {
      name,
      setKey(v) { key = String(v || '').trim(); share = ''; visitor = ''; },
      clearKey() { key = ''; share = ''; visitor = ''; },
      hasKey() { return !!(key || share); },
      // V6.22 : jeton du lien + jeton personnel du destinataire identifié.
      setShare(v, who) { share = String(v || '').trim(); visitor = String(who || '').trim(); key = ''; },
      isShare() { return !!share; },
      preparePopup, ensure, request, destroy, closePopupSoon, navigatePopup,
      get popupBlocked() { return st.popupBlocked; },
      get kind() { return st.kind; }
    };
  }

  window.NEWOSB_BRIDGE = { create, splitUrl, isWithin, TRUSTED_BRIDGE_ORIGIN, AuthError };
})();
