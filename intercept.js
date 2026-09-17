// intercept.js — MAIN-world content script on *.sharepoint.com (document_start, all_frames).
//
// Monkey-patches window.fetch and XMLHttpRequest to observe SharePoint Stream
// transcript metadata, then postMessage's it to the isolated-world content
// script running in the same frame.
//
// Transparency guarantees:
//   - fetch: always returns the ORIGINAL Response to the caller. We observe
//     via response.clone() on a detached promise chain that never feeds back.
//   - XHR: only adds a 'load' listener via addEventListener (non-destructive),
//     never overwrites onreadystatechange/onload, and always calls the
//     original open/send with the page's own `this` and arguments.
//
// Message shapes posted to window.location.origin:
//   { type: 'TTC_TRANSCRIPT_METADATA', data: any, url: string }
//   { type: 'TTC_BEARER_TOKEN',        token: string }

(function () {
  'use strict';

  // ---------- Pure helpers (exported for Node-based unit tests) ----------

  // URL filter for transcript metadata responses. Skips the binary body under
  // /content and the CDN media mirror under /cdnmedia/.
  function isTranscriptMetadataUrl(url) {
    if (typeof url !== 'string') return false;
    return (
      url.includes('transcripts') &&
      !url.includes('/content') &&
      !url.includes('/cdnmedia/')
    );
  }

  // True when a URL is a SharePoint REST call we want to lift tokens from.
  function isApiUrl(url) {
    return typeof url === 'string' && url.includes('/_api/v');
  }

  // Pull an Authorization bearer token out of a Headers instance, a plain
  // object, or the lower-cased fallback key. Returns null when absent.
  function extractAuthHeader(h) {
    if (!h) return null;
    if (typeof Headers !== 'undefined' && h instanceof Headers) {
      return h.get('Authorization') || h.get('authorization');
    }
    if (typeof h === 'object') {
      return h.Authorization || h.authorization || null;
    }
    return null;
  }

  // Given a Request|string and optional init, find the most recent
  // Authorization header. Precedence: the caller's `init.headers` override any
  // headers carried on a Request object (fetch spec behavior).
  function getAuthToken(resource, init) {
    if (init && init.headers) {
      const t = extractAuthHeader(init.headers);
      if (t) return t;
    }
    if (typeof Request !== 'undefined' && resource instanceof Request) {
      const t = extractAuthHeader(resource.headers);
      if (t) return t;
    }
    return null;
  }

  // Normalize fetch's first argument to a URL string.
  function urlOfResource(resource) {
    if (typeof resource === 'string') return resource;
    if (resource && typeof resource.url === 'string') return resource.url;
    return '';
  }

  // Post a payload to the same origin. Wrapped in try/catch so a hostile or
  // torn-down page context never breaks the underlying fetch/XHR call.
  function post(payload) {
    try {
      window.postMessage(payload, window.location.origin);
    } catch (err) {
      // swallow — never leak interception errors into page code
    }
  }

  // Expose helpers for page-side inspection and (in Node) for unit tests.
  // The window.__TTC handle is also used by synthetic test pages.
  const api = {
    isTranscriptMetadataUrl: isTranscriptMetadataUrl,
    isApiUrl: isApiUrl,
    extractAuthHeader: extractAuthHeader,
    getAuthToken: getAuthToken,
    urlOfResource: urlOfResource,
  };
  if (typeof window !== 'undefined') {
    try {
      window.__TTC = api;
    } catch (err) {
      // ignore
    }
  }
  if (typeof module !== 'undefined' && typeof module.exports !== 'undefined') {
    module.exports = api;
  }
  if (typeof window === 'undefined') return; // Node test path: helpers only

  // ---------- fetch hook ----------

  if (typeof window.fetch === 'function') {
    const originalFetch = window.fetch;

    window.fetch = function (resource, init) {
      // Call the page's fetch with the page's own `this` (window or a wrapper
      // some other script installed). Never await before returning — that
      // keeps microtask ordering identical to a native call.
      const result = originalFetch.apply(this, arguments);

      const url = urlOfResource(resource);

      if (isApiUrl(url)) {
        try {
          const token = getAuthToken(resource, init);
          if (token) {
            post({ type: 'TTC_BEARER_TOKEN', token: token });
          }
        } catch (err) {
          // swallow
        }
      }

      if (isTranscriptMetadataUrl(url)) {
        // Detached observation chain. Failures here MUST NOT affect the
        // caller's promise, so every step is guarded.
        result
          .then(function (response) {
            if (!response || typeof response.clone !== 'function') return;
            return response
              .clone()
              .json()
              .then(function (data) {
                post({
                  type: 'TTC_TRANSCRIPT_METADATA',
                  data: data,
                  url: url,
                });
              });
          })
          .catch(function () {
            // non-JSON body, aborted request, CORS — ignore all of it
          });
      }

      return result;
    };

    // Preserve identity markers some libraries check.
    try {
      Object.defineProperty(window.fetch, 'name', { value: 'fetch' });
      window.fetch.toString = function () {
        return 'function fetch() { [native code] }';
      };
    } catch (err) {
      // non-fatal
    }
  }

  // ---------- XMLHttpRequest hook ----------

  if (typeof XMLHttpRequest !== 'undefined' && XMLHttpRequest.prototype) {
    const proto = XMLHttpRequest.prototype;
    const originalOpen = proto.open;
    const originalSend = proto.send;
    const originalSetRequestHeader = proto.setRequestHeader;

    proto.open = function (method, url) {
      try {
        this.__ttcMethod = method;
        this.__ttcUrl = typeof url === 'string' ? url : String(url || '');
        this.__ttcAuth = null;
      } catch (err) {
        // swallow
      }
      return originalOpen.apply(this, arguments);
    };

    proto.setRequestHeader = function (name, value) {
      try {
        if (
          typeof name === 'string' &&
          name.toLowerCase() === 'authorization' &&
          typeof this.__ttcUrl === 'string' &&
          isApiUrl(this.__ttcUrl)
        ) {
          this.__ttcAuth = value;
        }
      } catch (err) {
        // swallow
      }
      return originalSetRequestHeader.apply(this, arguments);
    };

    proto.send = function () {
      try {
        const url = this.__ttcUrl || '';

        if (this.__ttcAuth && isApiUrl(url)) {
          post({ type: 'TTC_BEARER_TOKEN', token: this.__ttcAuth });
        }

        if (isTranscriptMetadataUrl(url)) {
          const xhr = this;
          // addEventListener (not .onload = ...) so the page's own handlers
          // are never displaced.
          xhr.addEventListener('load', function () {
            try {
              // responseType '' or 'text' keeps responseText readable; the
              // responseURL fallback covers redirects.
              const finalUrl = xhr.responseURL || url;
              const text = xhr.responseText;
              if (typeof text !== 'string' || !text) return;
              const data = JSON.parse(text);
              post({
                type: 'TTC_TRANSCRIPT_METADATA',
                data: data,
                url: finalUrl,
              });
            } catch (err) {
              // non-JSON body or read error — ignore
            }
          });
        }
      } catch (err) {
        // swallow — never break the page's XHR
      }
      return originalSend.apply(this, arguments);
    };
  }
})();
