// node_modules/@azure/msal-browser/dist/popup_relay/constants.mjs
var POPUP_RELAY_RESPONSE_TYPE = "msal:popup-relay-response:v1";

// node_modules/@azure/msal-common/dist-browser/error/AuthError.mjs
function getDefaultErrorMessage(code) {
  return `See https://aka.ms/msal.js.errors#${code} for details`;
}
var AuthError = class _AuthError extends Error {
  constructor(errorCode, correlationId, errorMessage, suberror) {
    const message = errorMessage || (errorCode ? getDefaultErrorMessage(errorCode) : "");
    const errorString = message ? `${errorCode}: ${message}` : errorCode;
    super(errorString);
    Object.setPrototypeOf(this, _AuthError.prototype);
    this.errorCode = errorCode || "";
    this.errorMessage = message || "";
    this.subError = suberror || "";
    this.correlationId = correlationId;
    this.name = "AuthError";
  }
};

// node_modules/@azure/msal-browser/dist/error/BrowserAuthError.mjs
function getDefaultErrorMessage2(code) {
  return `See https://aka.ms/msal.js.errors#${code} for details`;
}
var BrowserAuthError = class _BrowserAuthError extends AuthError {
  constructor(errorCode, correlationId, subError) {
    super(errorCode, correlationId, getDefaultErrorMessage2(errorCode), subError);
    Object.setPrototypeOf(this, _BrowserAuthError.prototype);
    this.name = "BrowserAuthError";
  }
};
function createBrowserAuthError(errorCode, correlationId, subError) {
  return new BrowserAuthError(errorCode, correlationId, subError);
}

// node_modules/@azure/msal-browser/dist/error/BrowserAuthErrorCodes.mjs
var popupWindowError = "popup_window_error";
var userCancelled = "user_cancelled";
var popupRelayUnsupportedFlow = "popup_relay_unsupported_flow";
var timedOut = "timed_out";

// node_modules/@azure/msal-browser/dist/protocol/Authorize.mjs
function createForm(frame, action, fields) {
  const form = frame.createElement("form");
  form.method = "post";
  form.action = action;
  Object.keys(fields).forEach((name) => {
    const input = frame.createElement("input");
    input.hidden = true;
    input.name = name;
    input.value = fields[name];
    form.appendChild(input);
  });
  frame.body.appendChild(form);
  return form;
}

// node_modules/@azure/msal-browser/dist/popup_relay/index.mjs
var DEFAULT_POPUP_RELAY_TIMEOUT_MS = 3e5;
var DEFAULT_POPUP_RELAY_WIDTH = 520;
var DEFAULT_POPUP_RELAY_HEIGHT = 640;
var BAD_REQUEST_SUBERROR = "popup_relay_bad_request";
var UNSAFE_URL_SUBERROR = "popup_relay_unsafe_url";
var UNTRUSTED_AUTHORITY_SUBERROR = "popup_relay_untrusted_authority";
var INVALID_ALLOWED_ORIGIN_SUBERROR = "popup_relay_invalid_allowed_origin";
function throwRelayError(subError) {
  throw createBrowserAuthError(popupRelayUnsupportedFlow, "", subError);
}
function tryParseUrl(raw) {
  try {
    return new URL(raw);
  } catch (e) {
    return null;
  }
}
function normalizeAllowedOrigins(allowedAuthorityOrigins) {
  return allowedAuthorityOrigins.map((entry) => {
    const parsed = typeof entry === "string" ? tryParseUrl(entry) : null;
    if (!parsed || parsed.protocol !== "https:" || parsed.username || parsed.password) {
      throwRelayError(INVALID_ALLOWED_ORIGIN_SUBERROR);
    }
    return parsed.origin;
  });
}
function validateNavigationTarget(raw, allowedAuthorityOrigins) {
  if (typeof raw !== "string" || !raw) {
    throwRelayError(BAD_REQUEST_SUBERROR);
  }
  const parsed = tryParseUrl(raw);
  if (!parsed || parsed.protocol !== "https:" || parsed.username || parsed.password) {
    throwRelayError(UNSAFE_URL_SUBERROR);
  }
  if (allowedAuthorityOrigins !== void 0 && !Array.isArray(allowedAuthorityOrigins)) {
    throwRelayError(INVALID_ALLOWED_ORIGIN_SUBERROR);
  }
  const normalizedAllowedOrigins = allowedAuthorityOrigins === void 0 ? void 0 : normalizeAllowedOrigins(allowedAuthorityOrigins);
  if (normalizedAllowedOrigins && normalizedAllowedOrigins.indexOf(parsed.origin) < 0) {
    throwRelayError(UNTRUSTED_AUTHORITY_SUBERROR);
  }
  return raw;
}
function parseRelayRequest(rawReq, allowedAuthorityOrigins) {
  let parsed;
  try {
    parsed = JSON.parse(rawReq);
  } catch (e) {
    parsed = null;
  }
  if (typeof parsed !== "object" || parsed === null) {
    throwRelayError(BAD_REQUEST_SUBERROR);
  }
  const request = parsed;
  const id = request.id;
  if (typeof id !== "string" || !id) {
    throwRelayError(BAD_REQUEST_SUBERROR);
  }
  if (request.method === "GET") {
    return {
      id,
      method: "GET",
      url: validateNavigationTarget(request.url, allowedAuthorityOrigins)
    };
  }
  if (request.method === "POST") {
    const action = validateNavigationTarget(request.action, allowedAuthorityOrigins);
    const rawFields = request.fields;
    if (typeof rawFields !== "object" || rawFields === null || Array.isArray(rawFields)) {
      throwRelayError(BAD_REQUEST_SUBERROR);
    }
    const fields = {};
    Object.keys(rawFields).forEach((name) => {
      const value = rawFields[name];
      if (typeof value !== "string") {
        throwRelayError(BAD_REQUEST_SUBERROR);
      }
      fields[name] = value;
    });
    return { id, method: "POST", action, fields };
  }
  return throwRelayError(BAD_REQUEST_SUBERROR);
}
function buildChildPopupFeatures(attributes) {
  const width = attributes?.popupSize?.width || DEFAULT_POPUP_RELAY_WIDTH;
  const height = attributes?.popupSize?.height || DEFAULT_POPUP_RELAY_HEIGHT;
  const features = [`width=${width}`, `height=${height}`];
  if (attributes?.popupPosition?.top !== void 0) {
    features.push(`top=${attributes.popupPosition.top}`);
  }
  if (attributes?.popupPosition?.left !== void 0) {
    features.push(`left=${attributes.popupPosition.left}`);
  }
  features.push("scrollbars=yes");
  return features.join(",");
}
function runPopupRelay(options) {
  const opener = window.opener;
  const targetOrigin = window.location.origin;
  const hashParams = new URLSearchParams(window.location.hash.replace(/^#/, ""));
  const rawReq = hashParams.get("req") || "";
  if (typeof window.history.replaceState === "function") {
    window.history.replaceState(null, "", window.location.origin + window.location.pathname);
  }
  if (!opener) {
    throw createBrowserAuthError(popupRelayUnsupportedFlow, "", "popup_relay_no_opener");
  }
  const request = parseRelayRequest(rawReq, options?.allowedAuthorityOrigins);
  const id = request.id;
  const channel = new BroadcastChannel(id);
  let settled = false;
  const timers = {};
  const relay = (message) => {
    if (settled) {
      return;
    }
    settled = true;
    try {
      channel.close();
    } catch (e) {
    }
    clearInterval(timers.closedPoll);
    clearTimeout(timers.timeoutId);
    opener.postMessage({ type: POPUP_RELAY_RESPONSE_TYPE, id, ...message }, targetOrigin);
    setTimeout(() => {
      try {
        window.close();
      } catch (e) {
      }
    }, 0);
  };
  channel.onmessage = (event) => {
    const payload = event.data && event.data.payload;
    if (payload) {
      relay({ payload });
    }
  };
  const features = buildChildPopupFeatures(options?.popupWindowAttributes);
  const childPopup = request.method === "POST" ? window.open("about:blank", "msalPopupRelayChild", features) : window.open(request.url, "msalPopupRelayChild", features);
  if (!childPopup) {
    relay({ error: popupWindowError });
    return;
  }
  if (request.method === "POST") {
    createForm(childPopup.document, request.action, request.fields).submit();
  }
  timers.closedPoll = window.setInterval(() => {
    if (childPopup.closed && !settled) {
      clearTimeout(timers.timeoutId);
      relay({ error: userCancelled });
    }
  }, 500);
  timers.timeoutId = window.setTimeout(() => {
    if (!settled) {
      clearInterval(timers.closedPoll);
      try {
        childPopup.close();
      } catch (e) {
      }
      relay({ error: timedOut });
    }
  }, options?.timeoutMs || DEFAULT_POPUP_RELAY_TIMEOUT_MS);
}

// relay.js
var button = document.querySelector("button");
var status = document.querySelector("[role=status]");
button.onclick = async () => {
  button.disabled = true;
  try {
    await runPopupRelay({ allowedAuthorityOrigins: ["https://login.microsoftonline.com"] });
  } catch (error) {
    status.textContent = "Microsoft 365 \xB7 " + (error.errorCode || error.name);
    button.disabled = false;
  }
};
/*! Bundled license information:

@azure/msal-browser/dist/popup_relay/constants.mjs:
@azure/msal-browser/dist/error/BrowserAuthError.mjs:
@azure/msal-browser/dist/error/BrowserAuthErrorCodes.mjs:
@azure/msal-browser/dist/protocol/Authorize.mjs:
@azure/msal-browser/dist/popup_relay/index.mjs:
  (*! @azure/msal-browser v5.29.0 2026-11-06 *)

@azure/msal-common/dist-browser/error/AuthError.mjs:
@azure/msal-common/dist-browser/index-browser.mjs:
  (*! @azure/msal-common v16.16.1 2026-11-06 *)
*/
