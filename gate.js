"use strict";

// Unlocks the password-protected site. Each page is published encrypted (site/<page>.enc); keys.json holds the site key
// locked by each password (PBKDF2-SHA256). The password is never stored: once it unlocks the site key, that key is kept
// for this browser tab (or on this computer if asked) and used to decrypt the page and its scripts and styles here.
(() => {
  const subtle = crypto.subtle;
  const page = decodeURIComponent(location.pathname.split("/").pop() || "") || "index.html";
  const libraries = new Set(["jszip.min.js", "pdfjs-display.js", "pdfjs-worker.js", "xlsx.full.min.js"]);
  const fromBase64 = text => Uint8Array.from(atob(text), c => c.charCodeAt(0));
  const toBase64 = bytes => btoa(String.fromCharCode(...bytes));
  const element = id => document.getElementById(id);

  const keysFor = async raw => ({
    enc: await subtle.importKey("raw", raw.slice(0, 32), "AES-CBC", false, ["decrypt"]),
    mac: await subtle.importKey("raw", raw.slice(32, 64), { name: "HMAC", hash: "SHA-256" }, false, ["verify"]),
  });
  // IV | HMAC-SHA256(IV | ciphertext) | AES-256-CBC ciphertext
  const openSealed = async (keys, bytes) => {
    const iv = bytes.slice(0, 16), tag = bytes.slice(16, 48), body = bytes.slice(48);
    const signed = new Uint8Array(16 + body.length);
    signed.set(iv);
    signed.set(body, 16);
    if (!await subtle.verify("HMAC", keys.mac, tag, signed)) throw new Error("locked");
    return new Uint8Array(await subtle.decrypt({ name: "AES-CBC", iv }, keys.enc, body));
  };
  const fetchBytes = async path => {
    const response = await fetch(path, { cache: "no-cache" });
    if (!response.ok) throw new Error("The site file " + path + " is missing.");
    return new Uint8Array(await response.arrayBuffer());
  };
  const stored = () => { try { return sessionStorage.getItem("site-key") || localStorage.getItem("site-key"); } catch { return null; } };
  const forget = () => { try { sessionStorage.removeItem("site-key"); localStorage.removeItem("site-key"); } catch { /* storage blocked */ } };

  async function unlock(password) {
    const config = await (await fetch("keys.json", { cache: "no-cache" })).json();
    const material = await subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveBits"]);
    for (const slot of config.slots) {
      const bits = new Uint8Array(await subtle.deriveBits({ name: "PBKDF2", hash: "SHA-256", salt: fromBase64(slot.salt), iterations: config.iterations }, material, 512));
      try { return await openSealed(await keysFor(bits), fromBase64(slot.data)); } catch { /* try the next password slot */ }
    }
    throw new Error("That password doesn't unlock the site.");
  }

  async function show(raw) {
    const keys = await keysFor(raw);
    const html = new TextDecoder().decode(await openSealed(keys, await fetchBytes("site/" + page + ".enc")));
    const doc = new DOMParser().parseFromString(html, "text/html");
    doc.querySelector('meta[http-equiv="Content-Security-Policy"]')?.remove();
    const decrypted = async (name, type) => URL.createObjectURL(new Blob([await openSealed(keys, await fetchBytes("site/" + name + ".enc"))], { type }));
    const own = path => path && !/^(?:[a-z]+:|\/\/|data:|#)/i.test(path) && !libraries.has(path);
    await Promise.all([...doc.querySelectorAll('link[rel="stylesheet"]')].map(async link => {
      if (own(link.getAttribute("href"))) link.setAttribute("href", await decrypted(link.getAttribute("href"), "text/css"));
    }));
    await Promise.all([...doc.querySelectorAll("img[src]")].map(async image => {
      const src = image.getAttribute("src");
      if (own(src)) image.setAttribute("src", await decrypted(src, /\.svg$/i.test(src) ? "image/svg+xml" : "image/png"));
    }));
    const sources = [...doc.querySelectorAll("script[src]")].map(script => script.getAttribute("src"));
    doc.querySelectorAll("script").forEach(script => script.remove());
    const urls = await Promise.all(sources.map(src => own(src) ? decrypted(src, "text/javascript") : src));
    document.title = doc.title;
    document.head.replaceChildren(...[...doc.head.childNodes].map(node => document.importNode(node, true)));
    document.body.replaceWith(document.importNode(doc.body, true));
    // Scripts run in order, as the page's deferred scripts would, then the page's start-up events fire.
    for (const url of urls) {
      const script = document.createElement("script");
      script.src = url;
      script.async = false;
      const loaded = new Promise((resolve, reject) => { script.onload = resolve; script.onerror = () => reject(new Error("A page script could not load.")); });
      document.head.append(script);
      await loaded;
    }
    document.dispatchEvent(new Event("DOMContentLoaded"));
    window.dispatchEvent(new Event("load"));
    // A Lock link in the menu forgets the key on this computer.
    const nav = document.querySelector("nav");
    if (nav) {
      const lock = Object.assign(document.createElement("a"), { href: "#", textContent: "Lock", className: "lock-link" });
      lock.addEventListener("click", event => { event.preventDefault(); forget(); location.reload(); });
      nav.append(lock);
    }
  }

  async function start() {
    const saved = stored();
    if (saved) {
      try { await show(fromBase64(saved)); return; } catch { forget(); }
    }
    element("gate-loading").hidden = true;
    element("gate").hidden = false;
    element("gate-password").focus();
    element("gate-form").addEventListener("submit", async event => {
      event.preventDefault();
      const button = element("gate-submit"), status = element("gate-status");
      button.disabled = true;
      status.textContent = "";
      try {
        const raw = await unlock(element("gate-password").value);
        const text = toBase64(raw);
        try {
          sessionStorage.setItem("site-key", text);
          if (element("gate-remember").checked) localStorage.setItem("site-key", text);
        } catch { /* storage blocked: unlock for this page only */ }
        await show(raw);
      } catch (error) {
        status.textContent = error.message;
        button.disabled = false;
      }
    });
  }

  start().catch(error => { element("gate-loading").textContent = error.message; });
})();
