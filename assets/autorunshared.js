/**
 * Auto-run Signature Insertion
 *
 * Two code paths:
 *   - TEST account (gabriel.williams@lilly.com): New optimized version with
 *     OfficeRuntime.storage (desktop) + sessionStorage (web) caching, XHR-based.
 *   - Everyone else: Proven production code (async/fetch) that works on OWA.
 */

// Production API URL — the API server lives in a separate repo
var API_BASE_URL = 'https://lilly-signature-addin.dc.lilly.com';

// Test account for canary rollout
var TEST_EMAIL = 'gabriell.williams@lilly.com';

// Cache TTL for the new code path (10 minutes)
var CACHE_TTL_MS = 10 * 60 * 1000;

console.log('Autorun signature handler loaded, API_BASE_URL:', API_BASE_URL);

// Lightweight ping to server for cache hit logging (fire-and-forget)
function logCacheHit(userEmail, cacheType, ageSeconds) {
  try {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', API_BASE_URL + '/api/cache-ping?email=' + encodeURIComponent(userEmail) + 
      '&type=' + cacheType + '&age=' + ageSeconds, true);
    xhr.send();
  } catch (e) { /* ignore errors - this is just for logging */ }
}

// ============================================================================
// ROUTER — picks which code path based on user email
// ============================================================================
function checkSignature(eventObj) {
  console.log('=== AUTO-SIGNATURE TRIGGERED ===');
  var userEmail = Office.context.mailbox.userProfile.emailAddress;
  console.log('User email:', userEmail);

  if (userEmail.toLowerCase() === TEST_EMAIL) {
    console.log('>>> TEST ACCOUNT — using legacy async/fetch code path');
    legacyCheckSignature(userEmail, eventObj);
  } else {
    console.log('>>> PRODUCTION — using new optimized XHR code path');
    newCheckSignature(userEmail, eventObj);
  }
}

// ============================================================================
// NEW CODE PATH — XHR + caching (for test account only)
// Works on desktop (no async/fetch) AND web (sessionStorage fallback)
// ============================================================================
function newCheckSignature(userEmail, eventObj) {
  var storageAvailable = typeof OfficeRuntime !== 'undefined' &&
    OfficeRuntime.storage &&
    typeof OfficeRuntime.storage.getItem === 'function';
  console.log('OfficeRuntime.storage available:', storageAvailable);

  if (!storageAvailable) {
    var webCached = readWebCache();
    if (webCached) {
      console.log('sessionStorage cache hit — using cached signature');
      logCacheHit(userEmail, 'sessionStorage', 0);
      newSetSignature(webCached, eventObj);
      return;
    }
    console.log('No cache — fetching from API');
    newFetchSignature(userEmail, eventObj);
    return;
  }

  OfficeRuntime.storage.getItem("cachedSignature")
    .then(function (cachedValue) {
      if (cachedValue) {
        try {
          var cached = JSON.parse(cachedValue);
          var age = Date.now() - cached.timestamp;
          if (age < CACHE_TTL_MS && cached.signatureHTML) {
            console.log('OfficeRuntime cache hit — age:', Math.round(age / 1000) + 's');
            logCacheHit(userEmail, 'OfficeRuntime', Math.round(age / 1000));
            newSetSignature(cached.signatureHTML, eventObj);
            return;
          }
        } catch (e) { /* bad cache — fall through */ }
      }
      console.log('Cache miss or expired — fetching from API');
      newFetchSignature(userEmail, eventObj);
    })
    .catch(function (err) {
      console.warn('Storage read failed:', err);
      newFetchSignature(userEmail, eventObj);
    });
}

function newFetchSignature(userEmail, eventObj) {
  var url = API_BASE_URL + "/signature?email=" + encodeURIComponent(userEmail);
  console.log('Fetching signature from:', url);
  var xhr = new XMLHttpRequest();
  xhr.open("GET", url, true);
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if (xhr.status === 200) {
        try {
          var data = JSON.parse(xhr.responseText);
          var cacheEntry = JSON.stringify({
            signatureHTML: data.signatureHTML,
            timestamp: Date.now()
          });
          var canCache = typeof OfficeRuntime !== 'undefined' &&
            OfficeRuntime.storage &&
            typeof OfficeRuntime.storage.setItem === 'function';
          if (canCache) {
            OfficeRuntime.storage.setItem("cachedSignature", cacheEntry)
              .then(function () {
                console.log('Signature cached in OfficeRuntime.storage');
                newSetSignature(data.signatureHTML, eventObj);
              })
              .catch(function (err) {
                console.warn('Cache write failed:', err);
                newSetSignature(data.signatureHTML, eventObj);
              });
          } else {
            writeWebCache(data.signatureHTML);
            newSetSignature(data.signatureHTML, eventObj);
          }
        } catch (e) {
          console.error('JSON parse failed:', e);
          eventObj.completed();
        }
      } else {
        console.error('API returned status:', xhr.status);
        eventObj.completed();
      }
    }
  };
  xhr.send();
}

function readWebCache() {
  try {
    if (typeof sessionStorage === 'undefined') return null;
    var raw = sessionStorage.getItem('cachedSignature');
    if (!raw) return null;
    var cached = JSON.parse(raw);
    var age = Date.now() - cached.timestamp;
    if (age < CACHE_TTL_MS && cached.signatureHTML) {
      console.log('sessionStorage cache — age:', Math.round(age / 1000) + 's');
      return cached.signatureHTML;
    }
    return null;
  } catch (e) { return null; }
}

function writeWebCache(signatureHTML) {
  try {
    if (typeof sessionStorage === 'undefined') return;
    sessionStorage.setItem('cachedSignature', JSON.stringify({
      signatureHTML: signatureHTML,
      timestamp: Date.now()
    }));
    console.log('Signature cached in sessionStorage');
  } catch (e) {
    console.warn('sessionStorage write failed:', e);
  }
}

function newSetSignature(signatureHTML, eventObj) {
  console.log('Inserting signature (NEW path), HTML length:', signatureHTML.length);

  // Disable Outlook's built-in client signature first, then set ours
  disableClientSignatureThen(function () {
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureHTML,
      { coercionType: "html", asyncContext: eventObj },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log('Signature inserted successfully (NEW path)');
        } else {
          console.error('Failed to insert signature:', asyncResult.error);
        }
        asyncResult.asyncContext.completed();
      }
    );
  });
}

// ============================================================================
// LEGACY CODE PATH — proven production code (async/fetch)
// This is what was working in production before. Not modified.
// ============================================================================
async function legacyCheckSignature(userEmail, eventObj) {
  try {
    var response = await fetch(API_BASE_URL + '/signature?email=' + encodeURIComponent(userEmail));

    if (!response.ok) {
      throw new Error('HTTP ' + response.status + ': ' + response.statusText);
    }

    var data = await response.json();
    console.log('✓ Signature fetched (LEGACY):', {
      source: data.source,
      generatedAt: data.generatedAt,
      htmlLength: data.signatureHTML ? data.signatureHTML.length : 0,
      authMethod: data.authMethod || 'legacy'
    });

    // Disable Outlook's built-in client signature first, then set ours
    disableClientSignatureThen(function () {
      Office.context.mailbox.item.body.setSignatureAsync(
        data.signatureHTML,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log('Signature inserted successfully (LEGACY)');
          } else {
            console.error('Failed to insert signature (LEGACY):', asyncResult.error);
          }
          eventObj.completed();
        }
      );
    });
  } catch (error) {
    console.error('Error in legacy signature flow:', error);
    eventObj.completed();
  }
}

// ============================================================================
// HELPER — disable Outlook's built-in client signature, then run callback
// If the API isn't available (older hosts), just proceeds without error.
// ============================================================================
function disableClientSignatureThen(callback) {
  if (Office.context.mailbox.item.disableClientSignatureAsync) {
    Office.context.mailbox.item.disableClientSignatureAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log('Client signature disabled');
      } else {
        console.warn('disableClientSignatureAsync failed:', asyncResult.error);
      }
      callback();
    });
  } else {
    console.log('disableClientSignatureAsync not available on this host — skipping');
    callback();
  }
}

// ============================================================================
// Quick Insert — called when user clicks the Quick Insert ribbon button.
// Registered here so autorunweb.html (Bouncer-whitelisted) serves as FunctionFile.
// ============================================================================
async function quickInsertSignature(event) {
  try {
    if (!Office.context?.mailbox?.item) {
      event.completed();
      return;
    }
    const userEmail = Office.context.mailbox.userProfile.emailAddress;
    const url = API_BASE_URL + '/signature?email=' + encodeURIComponent(userEmail);
    const response = await fetch(url);
    if (!response.ok) throw new Error('HTTP ' + response.status);
    const data = await response.json();
    Office.context.mailbox.item.body.setSignatureAsync(
      data.signatureHTML,
      { coercionType: 'html' },
      function(asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error('Quick insert failed:', asyncResult.error);
        }
        event.completed();
      }
    );
  } catch (error) {
    console.error('Quick insert error:', error);
    event.completed();
  }
}

// ============================================================================
// Register handlers — guard against double-registration.
// In OWA, autorunshared.js loads twice (via webpack bundle in autorunweb.html
// AND as the standalone JS override), causing a [DuplicatedName] warning in
// Office.js that invalidates the entire command surface and hides all buttons.
// ============================================================================
if (!window.__autorunHandlersRegistered) {
  window.__autorunHandlersRegistered = true;
  Office.actions.associate("checkSignature", checkSignature);
  Office.actions.associate("quickInsertSignature", quickInsertSignature);
}
